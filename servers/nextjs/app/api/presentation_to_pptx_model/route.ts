import { ApiError } from "@/models/errors";
import { NextRequest, NextResponse } from "next/server";
import puppeteer, { Browser, ElementHandle, Page } from "puppeteer";
import {
  ElementAttributes,
  SlideAttributesResult,
} from "@/types/element_attibutes";
import {
  convertElementAttributesToPptxSlides,
  mapToPptxFontName,
} from "@/utils/pptx_models_utils";
import { PptxPresentationModel } from "@/types/pptx_models";
import fs from "fs";
import path from "path";
import { v4 as uuidv4 } from "uuid";
import sharp from "sharp";

export const dynamic = "force-dynamic";
export const revalidate = 0;
const EXPORT_FONT_DEBUG = process.env.EXPORT_FONT_DEBUG === "1";

interface GetAllChildElementsAttributesArgs {
  element: ElementHandle<Element>;
  rootRect?: {
    left: number;
    top: number;
    width: number;
    height: number;
  } | null;
  depth?: number;
  inheritedFont?: ElementAttributes["font"];
  inheritedBackground?: ElementAttributes["background"];
  inheritedBorderRadius?: number[];
  inheritedZIndex?: number;
  inheritedOpacity?: number;
  domPath?: string;
  screenshotsDir: string;
}

export async function GET(request: NextRequest) {
  let browser: Browser | null = null;
  let page: Page | null = null;

  try {
    const id = await getPresentationId(request);
    [browser, page] = await getBrowserAndPage(id);
    await waitForExportReady(page);
    const screenshotsDir = getScreenshotsDir();

    if (EXPORT_FONT_DEBUG) {
      await logDomFontDiagnostics(page);
    }
    const { slides, speakerNotes } = await getSlidesAndSpeakerNotes(page);
    const slides_attributes = await getSlidesAttributes(slides, screenshotsDir);
    if (EXPORT_FONT_DEBUG) {
      logScrapedFontDiagnostics(slides_attributes);
    }
    await postProcessSlidesAttributes(
      slides_attributes,
      screenshotsDir,
      speakerNotes
    );
    await maybeWriteExportDebug(id, slides_attributes);
    const slides_pptx_models =
      convertElementAttributesToPptxSlides(slides_attributes);
    const presentation_pptx_model: PptxPresentationModel = {
      slides: slides_pptx_models,
    };

    await closeBrowserAndPage(browser, page);

    return NextResponse.json(presentation_pptx_model);
  } catch (error: any) {
    console.error(error);
    await closeBrowserAndPage(browser, page);
    if (error instanceof ApiError) {
      return NextResponse.json(error, { status: 400 });
    }
    return NextResponse.json(
      { detail: `Internal server error: ${error.message}` },
      { status: 500 }
    );
  }
}

async function logDomFontDiagnostics(page: Page) {
  try {
    const data = await page.evaluate(() => {
      const rootStyles = window.getComputedStyle(
        document.documentElement
      );
      const equipVar = rootStyles
        .getPropertyValue("--font-equip")
        .trim();
      const equipExtVar = rootStyles
        .getPropertyValue("--font-equip-ext")
        .trim();
      const nodes = Array.from(
        document.querySelectorAll(".ppt-title")
      );
      const sample = nodes.slice(0, 10).map((el) => {
        const cs = window.getComputedStyle(el);
        return {
          tagName: el.tagName.toLowerCase(),
          className: (el as HTMLElement).className || "",
          fontFamily: cs.fontFamily,
          fontWeight: cs.fontWeight,
          fontStyle: cs.fontStyle,
          outerHTML: (el as HTMLElement).outerHTML.slice(0, 200),
        };
      });
      return {
        equipVar,
        equipExtVar,
        pptTitleCount: nodes.length,
        pptTitleSample: sample,
      };
    });

    console.log("[PPTX_FONT_DEBUG] cssVars", {
      "--font-equip": data.equipVar,
      "--font-equip-ext": data.equipExtVar,
    });
    console.log("[PPTX_FONT_DEBUG] ppt-title count", data.pptTitleCount);
    for (const item of data.pptTitleSample) {
      console.log("[PPTX_FONT_DEBUG] ppt-title sample", item);
    }
  } catch (err) {
    console.warn("[PPTX_FONT_DEBUG] dom diagnostics failed", err);
  }
}

function logScrapedFontDiagnostics(
  slides_attributes: SlideAttributesResult[]
) {
  let totalTextElements = 0;
  let pptTitleElements = 0;

  for (const [slideIndex, slide] of slides_attributes.entries()) {
    for (const [elementIndex, el] of slide.elements.entries()) {
      if (!el.innerText || el.innerText.trim().length === 0) continue;
      totalTextElements += 1;

      const className = el.className || "";
      const hasPptTitle = className.includes("ppt-title");
      if (hasPptTitle) pptTitleElements += 1;

      const mapped = mapToPptxFontName(
        el.font?.name,
        el.font?.family,
        el.font?.weight,
        el.tagName
      );

      console.log("[PPTX_FONT_DEBUG] scraped-element", {
        slideIndex,
        elementIndex,
        elementId: el.id,
        dataEditableId: el.dataEditableId,
        tagName: el.tagName,
        className,
        fontName: el.font?.name,
        fontFamily: el.font?.family,
        fontWeight: el.font?.weight,
        fontStyle: el.font?.italic ? "italic" : "normal",
        mapped,
        hasPptTitle,
      });
    }
  }

  console.log("[PPTX_FONT_DEBUG] scraped summary", {
    totalTextElements,
    pptTitleElements,
  });
}

async function getPresentationId(request: NextRequest) {
  const id = request.nextUrl.searchParams.get("id");
  if (!id) {
    throw new ApiError("Presentation ID not found");
  }
  return id;
}

async function getBrowserAndPage(id: string): Promise<[Browser, Page]> {
  const browser = await puppeteer.launch({
    executablePath: process.env.PUPPETEER_EXECUTABLE_PATH,
    headless: true,
    args: [
      "--no-sandbox",
      "--disable-setuid-sandbox",
      "--disable-dev-shm-usage",
      "--disable-gpu",
      "--disable-web-security",
      "--disable-background-timer-throttling",
      "--disable-backgrounding-occluded-windows",
      "--disable-renderer-backgrounding",
      "--disable-features=TranslateUI",
      "--disable-ipc-flooding-protection",
    ],
  });

  const page = await browser.newPage();

  await page.setViewport({ width: 1280, height: 720, deviceScaleFactor: 1 });
  page.setDefaultNavigationTimeout(300000);
  page.setDefaultTimeout(300000);
  await page.goto(`http://localhost/pdf-maker?id=${id}`, {
    waitUntil: "networkidle0",
    timeout: 300000,
  });
  return [browser, page];
}

async function waitForExportReady(page: Page) {
  await page.addStyleTag({
    content: `*,*::before,*::after{animation:none !important;transition:none !important;}`,
  });
  await page.waitForFunction(
    () => (window as any).__PRESENTON_EXPORT_READY__ === true,
    { timeout: 60000 }
  );
}

async function maybeWriteExportDebug(
  presentationId: string,
  slidesAttributes: SlideAttributesResult[]
) {
  if (process.env.EXPORT_DEBUG !== "1") return;

  const debugDir = path.join(process.cwd(), "tmp", "export_debug");
  fs.mkdirSync(debugDir, { recursive: true });

  const runId = `${presentationId}-${Date.now()}`;
  const filePath = path.join(debugDir, `export-${runId}.json`);

  const debugPayload = {
    presentationId,
    slides: slidesAttributes.map((slide, slideIndex) => ({
      slideIndex,
      backgroundColor: slide.backgroundColor,
      elements: slide.elements.map((el, elementIndex) => ({
        index: elementIndex,
        domPath: el.domPath,
        depth: el.depth,
        zIndex: el.zIndex,
        bbox: el.position,
        text: el.innerText,
        elementId: el.id,
        className: el.className,
        tagName: el.tagName,
      })),
    })),
  };

  fs.writeFileSync(filePath, JSON.stringify(debugPayload, null, 2));
}

async function closeBrowserAndPage(browser: Browser | null, page: Page | null) {
  await page?.close();
  await browser?.close();
}

function getScreenshotsDir() {
  const tempDir = process.env.TEMP_DIRECTORY;
  if (!tempDir) {
    console.warn(
      "TEMP_DIRECTORY environment variable not set, skipping screenshot"
    );
    throw new ApiError("TEMP_DIRECTORY environment variable not set");
  }
  const screenshotsDir = path.join(tempDir, "screenshots");
  if (!fs.existsSync(screenshotsDir)) {
    fs.mkdirSync(screenshotsDir, { recursive: true });
  }
  return screenshotsDir;
}

async function postProcessSlidesAttributes(
  slidesAttributes: SlideAttributesResult[],
  screenshotsDir: string,
  speakerNotes: string[]
) {
  for (const [index, slideAttributes] of slidesAttributes.entries()) {
    for (const [elementIndex, element] of slideAttributes.elements.entries()) {
      if (element.should_screenshot) {
        const stableId = element.domPath
          ? `slide${index}-${element.domPath}`
          : `slide${index}-el${elementIndex}`;
        const screenshotPath = await screenshotElement(
          element,
          screenshotsDir,
          stableId
        );
        element.imageSrc = screenshotPath;
        element.should_screenshot = false;
        element.objectFit = "cover";
        element.element = undefined;
      }
    }
    slideAttributes.speakerNote = speakerNotes[index];
  }
}

async function screenshotElement(
  element: ElementAttributes,
  screenshotsDir: string,
  stableId?: string
) {
  const safeId = stableId
    ? stableId.replace(/[^a-zA-Z0-9._-]/g, "_")
    : uuidv4();
  const screenshotPath = path.join(
    screenshotsDir,
    `${safeId}.png`
  ) as `${string}.png`;

  // For SVG elements, use convertSvgToPng
  if (element.tagName === "svg") {
    const pngBuffer = await convertSvgToPng(element);
    fs.writeFileSync(screenshotPath, pngBuffer);
    return screenshotPath;
  }

  // Hide all elements except the target element and its ancestors
  await element.element?.evaluate(
    (el) => {
      const originalOpacities = new Map();

      const hideAllExcept = (targetElement: Element) => {
        const allElements = document.querySelectorAll("*");

        allElements.forEach((elem) => {
          const computedStyle = window.getComputedStyle(elem);
          originalOpacities.set(elem, computedStyle.opacity);

          if (
            targetElement === elem ||
            targetElement.contains(elem) ||
            elem.contains(targetElement)
          ) {
            (elem as HTMLElement).style.opacity = computedStyle.opacity || "1";
            return;
          }

          (elem as HTMLElement).style.opacity = "0";
        });
      };

      hideAllExcept(el);

      (el as any).__restoreStyles = () => {
        originalOpacities.forEach((opacity, elem) => {
          (elem as HTMLElement).style.opacity = opacity;
        });
      };
    },
    element.opacity,
    element.font?.color
  );

  const screenshot = await element.element?.screenshot({
    path: screenshotPath,
  });
  if (!screenshot) {
    throw new ApiError("Failed to screenshot element");
  }

  await element.element?.evaluate((el) => {
    if ((el as any).__restoreStyles) {
      (el as any).__restoreStyles();
    }
  });

  return screenshotPath;
}

const convertSvgToPng = async (element_attibutes: ElementAttributes) => {
  const svgHtml =
    (await element_attibutes.element?.evaluate((el) => {
      // Apply font color
      const fontColor = window.getComputedStyle(el).color;
      (el as HTMLElement).style.color = fontColor;

      return el.outerHTML;
    })) || "";

  const svgBuffer = Buffer.from(svgHtml);
  const pngBuffer = await sharp(svgBuffer)
    .resize(
      Math.round(element_attibutes.position!.width!),
      Math.round(element_attibutes.position!.height!)
    )
    .toFormat("png")
    .toBuffer();
  return pngBuffer;
};

async function getSlidesAttributes(
  slides: ElementHandle<Element>[],
  screenshotsDir: string
): Promise<SlideAttributesResult[]> {
  const slideAttributes = await Promise.all(
    slides.map((slide) =>
      getAllChildElementsAttributes({ element: slide, screenshotsDir })
    )
  );
  return slideAttributes;
}

async function getSlidesAndSpeakerNotes(page: Page) {
  const slides_wrapper = await getSlidesWrapper(page);
  const speakerNotes = await getSpeakerNotes(slides_wrapper);
  const slides = await slides_wrapper.$$(":scope > div > div");
  return { slides, speakerNotes };
}

async function getSlidesWrapper(page: Page): Promise<ElementHandle<Element>> {
  const slides_wrapper = await page.$("#presentation-slides-wrapper");
  if (!slides_wrapper) {
    throw new ApiError("Presentation slides not found");
  }
  return slides_wrapper;
}

async function getSpeakerNotes(slides_wrapper: ElementHandle<Element>) {
  return await slides_wrapper.evaluate((el) => {
    return Array.from(el.querySelectorAll("[data-speaker-note]")).map(
      (el) => el.getAttribute("data-speaker-note") || ""
    );
  });
}

async function getAllChildElementsAttributes({
  element,
  rootRect = null,
  depth = 0,
  inheritedFont,
  inheritedBackground,
  inheritedBorderRadius,
  inheritedZIndex,
  inheritedOpacity,
  domPath = "",
  screenshotsDir,
}: GetAllChildElementsAttributesArgs): Promise<SlideAttributesResult> {
  if (!rootRect) {
    const rootAttributes = await getElementAttributes(element);
    inheritedFont = rootAttributes.font;
    inheritedBackground = rootAttributes.background;
    inheritedZIndex = rootAttributes.zIndex;
    inheritedOpacity = rootAttributes.opacity;
    rootRect = {
      left: rootAttributes.position?.left ?? 0,
      top: rootAttributes.position?.top ?? 0,
      width: rootAttributes.position?.width ?? 1280,
      height: rootAttributes.position?.height ?? 720,
    };
  }

  const directChildElementHandles = await element.$$(":scope > *");

  const allResults: { attributes: ElementAttributes; depth: number }[] = [];

  for (const [childIndex, childElementHandle] of directChildElementHandles.entries()) {
    const childDomPath = domPath ? `${domPath}.${childIndex}` : `${childIndex}`;
    const attributes = await getElementAttributes(childElementHandle);
    attributes.domPath = childDomPath;
    attributes.depth = depth;

    if (
      ["style", "script", "link", "meta", "path"].includes(attributes.tagName)
    ) {
      continue;
    }

    if (
      inheritedFont &&
      !attributes.font &&
      attributes.innerText &&
      attributes.innerText.trim().length > 0
    ) {
      attributes.font = inheritedFont;
    }
    if (inheritedBackground && !attributes.background && attributes.shadow) {
      attributes.background = inheritedBackground;
    }
    if (inheritedBorderRadius && !attributes.borderRadius) {
      attributes.borderRadius = inheritedBorderRadius;
    }
    if (inheritedZIndex !== undefined && attributes.zIndex === 0) {
      attributes.zIndex = inheritedZIndex;
    }
    if (
      inheritedOpacity !== undefined &&
      (attributes.opacity === undefined || attributes.opacity === 1)
    ) {
      attributes.opacity = inheritedOpacity;
    }

    if (
      attributes.position &&
      attributes.position.left !== undefined &&
      attributes.position.top !== undefined
    ) {
      attributes.position = {
        left: attributes.position.left - rootRect!.left,
        top: attributes.position.top - rootRect!.top,
        width: attributes.position.width,
        height: attributes.position.height,
      };
    }

    // Ignore elements with no size (width or height)
    if (
      attributes.position === undefined ||
      attributes.position.width === undefined ||
      attributes.position.height === undefined ||
      attributes.position.width === 0 ||
      attributes.position.height === 0
    ) {
      continue;
    }

    // If element is paragraph and contains only inline formatting tags, don't go deeper
    if (attributes.tagName === "p") {
      const innerElementTagNames = await childElementHandle.evaluate((el) => {
        return Array.from(el.querySelectorAll("*")).map((e) =>
          e.tagName.toLowerCase()
        );
      });

      const allowedInlineTags = new Set(["strong", "u", "em", "code", "s"]);
      const hasOnlyAllowedInlineTags = innerElementTagNames.every((tag) =>
        allowedInlineTags.has(tag)
      );

      if (innerElementTagNames.length > 0 && hasOnlyAllowedInlineTags) {
        attributes.innerText = await childElementHandle.evaluate((el) => {
          return el.innerHTML;
        });
        allResults.push({ attributes, depth });
        continue;
      }
    }

    if (
      attributes.tagName === "svg" ||
      attributes.tagName === "canvas" ||
      attributes.tagName === "table"
    ) {
      attributes.should_screenshot = true;
      attributes.element = childElementHandle;
    }

    allResults.push({ attributes, depth });

    // If the element is a canvas, or table, we don't need to go deeper
    if (attributes.should_screenshot && attributes.tagName !== "svg") {
      continue;
    }

    const childResults = await getAllChildElementsAttributes({
      element: childElementHandle,
      rootRect: rootRect,
      depth: depth + 1,
      domPath: childDomPath,
      inheritedFont: attributes.font || inheritedFont,
      inheritedBackground: attributes.background || inheritedBackground,
      inheritedBorderRadius: attributes.borderRadius || inheritedBorderRadius,
      inheritedZIndex: attributes.zIndex || inheritedZIndex,
      inheritedOpacity: attributes.opacity || inheritedOpacity,
      screenshotsDir,
    });
    allResults.push(
      ...childResults.elements.map((attr) => ({
        attributes: {
          ...attr,
          depth: depth + 1,
        },
        depth: depth + 1,
      }))
    );
  }

  let backgroundColor = inheritedBackground?.color;
  if (depth === 0) {
    const elementsWithRootPosition = allResults.filter(({ attributes }) => {
      return (
        attributes.position &&
        attributes.position.left === 0 &&
        attributes.position.top === 0 &&
        attributes.position.width === rootRect!.width &&
        attributes.position.height === rootRect!.height
      );
    });

    for (const { attributes } of elementsWithRootPosition) {
      if (attributes.background && attributes.background.color) {
        backgroundColor = attributes.background.color;
        break;
      }
    }
  }

  const filteredResults =
    depth === 0
      ? allResults.filter(({ attributes }) => {
          const hasBackground =
            attributes.background && attributes.background.color;
          const hasBorder = attributes.border && attributes.border.color;
          const hasShadow = attributes.shadow && attributes.shadow.color;
          const hasText =
            attributes.innerText && attributes.innerText.trim().length > 0;
          const hasImage = attributes.imageSrc;
          const isSvg = attributes.tagName === "svg";
          const isCanvas = attributes.tagName === "canvas";
          const isTable = attributes.tagName === "table";

          const occupiesRoot =
            attributes.position &&
            attributes.position.left === 0 &&
            attributes.position.top === 0 &&
            attributes.position.width === rootRect!.width &&
            attributes.position.height === rootRect!.height;

          const hasVisualProperties =
            hasBackground || hasBorder || hasShadow || hasText;
          const hasSpecialContent = hasImage || isSvg || isCanvas || isTable;

          return (hasVisualProperties && !occupiesRoot) || hasSpecialContent;
        })
      : allResults;

  if (depth === 0) {
    const sortedElements = filteredResults
      .sort((a, b) => {
        const zIndexA = a.attributes.zIndex || 0;
        const zIndexB = b.attributes.zIndex || 0;
        const zIndexDiff = zIndexA - zIndexB;

        if (zIndexDiff !== 0) {
          return zIndexDiff;
        }

        const depthDiff = a.depth - b.depth;
        if (depthDiff !== 0) {
          return depthDiff;
        }

        const domPathA = a.attributes.domPath || "";
        const domPathB = b.attributes.domPath || "";
        if (domPathA !== domPathB) {
          return domPathA.localeCompare(domPathB);
        }

        const idA = a.attributes.id || "";
        const idB = b.attributes.id || "";
        if (idA !== idB) {
          return idA.localeCompare(idB);
        }

        return 0;
      })
      .map(({ attributes }) => {
        if (
          attributes.shadow &&
          attributes.shadow.color &&
          (!attributes.background || !attributes.background.color) &&
          backgroundColor
        ) {
          attributes.background = {
            color: backgroundColor,
            opacity: undefined,
          };
        }
        return attributes;
      });

    return {
      elements: sortedElements,
      backgroundColor,
    };
  } else {
    return {
      elements: filteredResults.map(({ attributes }) => attributes),
      backgroundColor,
    };
  }
}

async function getElementAttributes(
  element: ElementHandle<Element>
): Promise<ElementAttributes> {
  const attributes = await element.evaluate((el: Element) => {
    function colorToHex(color: string): {
      hex: string | undefined;
      opacity: number | undefined;
    } {
      if (!color || color === "transparent" || color === "rgba(0, 0, 0, 0)") {
        return { hex: undefined, opacity: undefined };
      }

      if (color.startsWith("rgba(") || color.startsWith("hsla(")) {
        const match = color.match(/rgba?\(([^)]+)\)|hsla?\(([^)]+)\)/);
        if (match) {
          const values = match[1] || match[2];
          const parts = values.split(",").map((part) => part.trim());

          if (parts.length >= 4) {
            const opacity = parseFloat(parts[3]);
            const rgbColor = color
              .replace(/rgba?\(|hsla?\(|\)/g, "")
              .split(",")
              .slice(0, 3)
              .join(",");
            const rgbString = color.startsWith("rgba")
              ? `rgb(${rgbColor})`
              : `hsl(${rgbColor})`;

            const canvas = document.createElement("canvas");
            const ctx = canvas.getContext("2d");
            if (ctx) {
              ctx.fillStyle = rgbString;
              const hexColor = ctx.fillStyle;
              const hex = hexColor.startsWith("#")
                ? hexColor.substring(1)
                : hexColor;
              const result = {
                hex,
                opacity: isNaN(opacity) ? undefined : opacity,
              };

              return result;
            }
          }
        }
      }

      if (color.startsWith("rgb(") || color.startsWith("hsl(")) {
        const canvas = document.createElement("canvas");
        const ctx = canvas.getContext("2d");
        if (ctx) {
          ctx.fillStyle = color;
          const hexColor = ctx.fillStyle;
          const hex = hexColor.startsWith("#")
            ? hexColor.substring(1)
            : hexColor;
          return { hex, opacity: undefined };
        }
      }

      if (color.startsWith("#")) {
        const hex = color.substring(1);
        return { hex, opacity: undefined };
      }

      const canvas = document.createElement("canvas");
      const ctx = canvas.getContext("2d");
      if (!ctx) return { hex: color, opacity: undefined };

      ctx.fillStyle = color;
      const hexColor = ctx.fillStyle;
      const hex = hexColor.startsWith("#") ? hexColor.substring(1) : hexColor;
      const result = { hex, opacity: undefined };

      return result;
    }

    function hasOnlyTextNodes(el: Element): boolean {
      const children = el.childNodes;
      for (let i = 0; i < children.length; i++) {
        const child = children[i];
        if (child.nodeType === Node.ELEMENT_NODE) {
          return false;
        }
      }
      return true;
    }

    function parsePosition(el: Element) {
      const rect = el.getBoundingClientRect();
      return {
        left: isFinite(rect.left) ? rect.left : 0,
        top: isFinite(rect.top) ? rect.top : 0,
        width: isFinite(rect.width) ? rect.width : 0,
        height: isFinite(rect.height) ? rect.height : 0,
      };
    }

    function parseBackground(computedStyles: CSSStyleDeclaration) {
      const backgroundColorResult = colorToHex(computedStyles.backgroundColor);

      const background = {
        color: backgroundColorResult.hex,
        opacity: backgroundColorResult.opacity,
      };

      // Return undefined if background has no meaningful values
      if (!background.color && background.opacity === undefined) {
        return undefined;
      }

      return background;
    }

    function parseBackgroundImage(computedStyles: CSSStyleDeclaration) {
      const backgroundImage = computedStyles.backgroundImage;

      if (!backgroundImage || backgroundImage === "none") {
        return undefined;
      }

      // Extract URL from background-image style
      const urlMatch = backgroundImage.match(/url\(['"]?([^'"]+)['"]?\)/);
      if (urlMatch && urlMatch[1]) {
        return urlMatch[1];
      }

      return undefined;
    }

    function parseBorder(computedStyles: CSSStyleDeclaration) {
      const borderColorResult = colorToHex(computedStyles.borderColor);
      const borderWidth = parseFloat(computedStyles.borderWidth);

      if (borderWidth === 0) {
        return undefined;
      }

      const border = {
        color: borderColorResult.hex,
        width: isNaN(borderWidth) ? undefined : borderWidth,
        opacity: borderColorResult.opacity,
      };

      // Return undefined if border has no meaningful values
      if (
        !border.color &&
        border.width === undefined &&
        border.opacity === undefined
      ) {
        return undefined;
      }

      return border;
    }

    function parseShadow(computedStyles: CSSStyleDeclaration) {
      const boxShadow = computedStyles.boxShadow;
      if (boxShadow !== "none") {
      }
      let shadow: {
        offset?: [number, number];
        color?: string;
        opacity?: number;
        radius?: number;
        angle?: number;
        spread?: number;
        inset?: boolean;
      } = {};

      if (boxShadow && boxShadow !== "none") {
        const shadows: string[] = [];
        let currentShadow = "";
        let parenCount = 0;

        for (let i = 0; i < boxShadow.length; i++) {
          const char = boxShadow[i];
          if (char === "(") {
            parenCount++;
          } else if (char === ")") {
            parenCount--;
          } else if (char === "," && parenCount === 0) {
            shadows.push(currentShadow.trim());
            currentShadow = "";
            continue;
          }
          currentShadow += char;
        }

        if (currentShadow.trim()) {
          shadows.push(currentShadow.trim());
        }

        let selectedShadow = "";
        let bestShadowScore = -1;

        for (let i = 0; i < shadows.length; i++) {
          const shadowStr = shadows[i];

          const shadowParts = shadowStr.split(" ");
          const numericParts: number[] = [];
          const colorParts: string[] = [];
          let isInset = false;
          let currentColor = "";
          let inColorFunction = false;

          for (let j = 0; j < shadowParts.length; j++) {
            const part = shadowParts[j];
            const trimmedPart = part.trim();
            if (trimmedPart === "") continue;

            if (trimmedPart.toLowerCase() === "inset") {
              isInset = true;
              continue;
            }

            if (trimmedPart.match(/^(rgba?|hsla?)\s*\(/i)) {
              inColorFunction = true;
              currentColor = trimmedPart;
              continue;
            }

            if (inColorFunction) {
              currentColor += " " + trimmedPart;

              const openParens = (currentColor.match(/\(/g) || []).length;
              const closeParens = (currentColor.match(/\)/g) || []).length;

              if (openParens <= closeParens) {
                colorParts.push(currentColor);
                currentColor = "";
                inColorFunction = false;
              }
              continue;
            }

            const numericValue = parseFloat(trimmedPart);
            if (!isNaN(numericValue)) {
              numericParts.push(numericValue);
            } else {
              colorParts.push(trimmedPart);
            }
          }

          let hasVisibleColor = false;
          if (colorParts.length > 0) {
            const shadowColor = colorParts.join(" ");
            const colorResult = colorToHex(shadowColor);
            hasVisibleColor = !!(
              colorResult.hex &&
              colorResult.hex !== "000000" &&
              colorResult.opacity !== 0
            );
          }

          const hasNonZeroValues = numericParts.some((value) => value !== 0);

          let shadowScore = 0;
          if (hasNonZeroValues) {
            shadowScore += numericParts.filter((value) => value !== 0).length;
          }
          if (hasVisibleColor) {
            shadowScore += 2;
          }

          if (
            (hasNonZeroValues || hasVisibleColor) &&
            shadowScore > bestShadowScore
          ) {
            selectedShadow = shadowStr;
            bestShadowScore = shadowScore;
          }
        }

        if (!selectedShadow && shadows.length > 0) {
          selectedShadow = shadows[0];
        }

        if (selectedShadow) {
          const shadowParts = selectedShadow.split(" ");
          const numericParts: number[] = [];
          const colorParts: string[] = [];
          let isInset = false;
          let currentColor = "";
          let inColorFunction = false;

          for (let i = 0; i < shadowParts.length; i++) {
            const part = shadowParts[i];
            const trimmedPart = part.trim();
            if (trimmedPart === "") continue;

            if (trimmedPart.toLowerCase() === "inset") {
              isInset = true;
              continue;
            }

            if (trimmedPart.match(/^(rgba?|hsla?)\s*\(/i)) {
              inColorFunction = true;
              currentColor = trimmedPart;
              continue;
            }

            if (inColorFunction) {
              currentColor += " " + trimmedPart;

              const openParens = (currentColor.match(/\(/g) || []).length;
              const closeParens = (currentColor.match(/\)/g) || []).length;

              if (openParens <= closeParens) {
                colorParts.push(currentColor);
                currentColor = "";
                inColorFunction = false;
              }
              continue;
            }

            const numericValue = parseFloat(trimmedPart);
            if (!isNaN(numericValue)) {
              numericParts.push(numericValue);
            } else {
              colorParts.push(trimmedPart);
            }
          }

          if (numericParts.length >= 2) {
            const offsetX = numericParts[0];
            const offsetY = numericParts[1];
            const blurRadius = numericParts.length >= 3 ? numericParts[2] : 0;
            const spreadRadius = numericParts.length >= 4 ? numericParts[3] : 0;

            // Only create shadow if color is present
            if (colorParts.length > 0) {
              const shadowColor = colorParts.join(" ");
              const shadowColorResult = colorToHex(shadowColor);

              if (shadowColorResult.hex) {
                shadow = {
                  offset: [offsetX, offsetY],
                  color: shadowColorResult.hex,
                  opacity: shadowColorResult.opacity,
                  radius: blurRadius,
                  spread: spreadRadius,
                  inset: isInset,
                  angle: Math.atan2(offsetY, offsetX) * (180 / Math.PI),
                };
              }
            }
          }
        }
      }

      // Return undefined if shadow is empty (no meaningful values)
      if (Object.keys(shadow).length === 0) {
        return undefined;
      }

      return shadow;
    }

    function parseFont(computedStyles: CSSStyleDeclaration) {
      const fontSize = parseFloat(computedStyles.fontSize);
      const fontWeight = parseInt(computedStyles.fontWeight);
      const fontColorResult = colorToHex(computedStyles.color);
      const fontFamily = computedStyles.fontFamily;
      const fontStyle = computedStyles.fontStyle;

      let fontName = undefined;
      if (fontFamily !== "initial") {
        const firstFont = fontFamily.split(",")[0].trim().replace(/['"]/g, "");
        fontName = firstFont;
      }

      const font = {
        name: fontName,
        family: fontFamily,
        size: isNaN(fontSize) ? undefined : fontSize,
        weight: isNaN(fontWeight) ? undefined : fontWeight,
        color: fontColorResult.hex,
        italic: fontStyle === "italic",
      };

      // Return undefined if font has no meaningful values
      if (
        !font.name &&
        font.size === undefined &&
        font.weight === undefined &&
        !font.color &&
        !font.italic
      ) {
        return undefined;
      }

      return font;
    }

    function parseLineHeight(computedStyles: CSSStyleDeclaration, el: Element) {
      const lineHeight = computedStyles.lineHeight;
      const innerText = el.textContent || "";

      const htmlEl = el as HTMLElement;

      const fontSize = parseFloat(computedStyles.fontSize);
      const computedLineHeight = parseFloat(computedStyles.lineHeight);

      const singleLineHeight = !isNaN(computedLineHeight)
        ? computedLineHeight
        : fontSize * 1.2;

      const hasExplicitLineBreaks =
        innerText.includes("\n") ||
        innerText.includes("\r") ||
        innerText.includes("\r\n");
      const hasTextWrapping = htmlEl.offsetHeight > singleLineHeight * 2;
      const hasOverflow = htmlEl.scrollHeight > htmlEl.clientHeight;

      const isMultiline =
        hasExplicitLineBreaks || hasTextWrapping || hasOverflow;

      if (isMultiline && lineHeight && lineHeight !== "normal") {
        const parsedLineHeight = parseFloat(lineHeight);
        if (!isNaN(parsedLineHeight)) {
          return parsedLineHeight;
        }
      }

      return undefined;
    }

    function parseMargin(computedStyles: CSSStyleDeclaration) {
      const marginTop = parseFloat(computedStyles.marginTop);
      const marginBottom = parseFloat(computedStyles.marginBottom);
      const marginLeft = parseFloat(computedStyles.marginLeft);
      const marginRight = parseFloat(computedStyles.marginRight);
      const marginObj = {
        top: isNaN(marginTop) ? undefined : marginTop,
        bottom: isNaN(marginBottom) ? undefined : marginBottom,
        left: isNaN(marginLeft) ? undefined : marginLeft,
        right: isNaN(marginRight) ? undefined : marginRight,
      };

      return marginObj.top === 0 &&
        marginObj.bottom === 0 &&
        marginObj.left === 0 &&
        marginObj.right === 0
        ? undefined
        : marginObj;
    }

    function parsePadding(computedStyles: CSSStyleDeclaration) {
      const paddingTop = parseFloat(computedStyles.paddingTop);
      const paddingBottom = parseFloat(computedStyles.paddingBottom);
      const paddingLeft = parseFloat(computedStyles.paddingLeft);
      const paddingRight = parseFloat(computedStyles.paddingRight);
      const paddingObj = {
        top: isNaN(paddingTop) ? undefined : paddingTop,
        bottom: isNaN(paddingBottom) ? undefined : paddingBottom,
        left: isNaN(paddingLeft) ? undefined : paddingLeft,
        right: isNaN(paddingRight) ? undefined : paddingRight,
      };

      return paddingObj.top === 0 &&
        paddingObj.bottom === 0 &&
        paddingObj.left === 0 &&
        paddingObj.right === 0
        ? undefined
        : paddingObj;
    }

    function parseBorderRadius(
      computedStyles: CSSStyleDeclaration,
      el: Element
    ) {
      const borderRadius = computedStyles.borderRadius;
      let borderRadiusValue;

      if (borderRadius && borderRadius !== "0px") {
        const radiusParts = borderRadius
          .split(" ")
          .map((part) => parseFloat(part));
        if (radiusParts.length === 1) {
          borderRadiusValue = [
            radiusParts[0],
            radiusParts[0],
            radiusParts[0],
            radiusParts[0],
          ];
        } else if (radiusParts.length === 2) {
          borderRadiusValue = [
            radiusParts[0],
            radiusParts[1],
            radiusParts[0],
            radiusParts[1],
          ];
        } else if (radiusParts.length === 3) {
          borderRadiusValue = [
            radiusParts[0],
            radiusParts[1],
            radiusParts[2],
            radiusParts[1],
          ];
        } else if (radiusParts.length === 4) {
          borderRadiusValue = radiusParts;
        }

        // Clamp border radius values to be between 0 and half the width/height
        if (borderRadiusValue) {
          const rect = el.getBoundingClientRect();
          const maxRadiusX = rect.width / 2;
          const maxRadiusY = rect.height / 2;

          borderRadiusValue = borderRadiusValue.map((radius, index) => {
            // For top-left and bottom-right corners, use maxRadiusX
            // For top-right and bottom-left corners, use maxRadiusY
            const maxRadius =
              index === 0 || index === 2 ? maxRadiusX : maxRadiusY;
            return Math.max(0, Math.min(radius, maxRadius));
          });
        }
      }

      return borderRadiusValue;
    }

    function parseShape(el: Element, borderRadiusValue: number[] | undefined) {
      if (el.tagName.toLowerCase() === "img") {
        return borderRadiusValue &&
          borderRadiusValue.length === 4 &&
          borderRadiusValue.every((radius: number) => radius === 50)
          ? "circle"
          : "rectangle";
      }
      return undefined;
    }

    function parseFilters(computedStyles: CSSStyleDeclaration) {
      const filter = computedStyles.filter;
      if (!filter || filter === "none") {
        return undefined;
      }

      const filters: {
        invert?: number;
        brightness?: number;
        contrast?: number;
        saturate?: number;
        hueRotate?: number;
        blur?: number;
        grayscale?: number;
        sepia?: number;
        opacity?: number;
      } = {};

      // Parse filter functions
      const filterFunctions = filter.match(/[a-zA-Z]+\([^)]*\)/g);
      if (filterFunctions) {
        filterFunctions.forEach((func) => {
          const match = func.match(/([a-zA-Z]+)\(([^)]*)\)/);
          if (match) {
            const filterType = match[1];
            const value = parseFloat(match[2]);

            if (!isNaN(value)) {
              switch (filterType) {
                case "invert":
                  filters.invert = value;
                  break;
                case "brightness":
                  filters.brightness = value;
                  break;
                case "contrast":
                  filters.contrast = value;
                  break;
                case "saturate":
                  filters.saturate = value;
                  break;
                case "hue-rotate":
                  filters.hueRotate = value;
                  break;
                case "blur":
                  filters.blur = value;
                  break;
                case "grayscale":
                  filters.grayscale = value;
                  break;
                case "sepia":
                  filters.sepia = value;
                  break;
                case "opacity":
                  filters.opacity = value;
                  break;
              }
            }
          }
        });
      }

      // Return undefined if no filters were parsed
      return Object.keys(filters).length > 0 ? filters : undefined;
    }

    function parseListInfo(el: Element) {
      const liEl = el.closest("li");
      if (!liEl) {
        return {
          isListItem: false,
          listType: undefined,
          listLevel: undefined,
          listStyleType: undefined,
          listStylePosition: undefined,
          listIndent: undefined,
          listItemIndex: undefined,
        };
      }

      const listEl = liEl.closest("ul,ol");
      const listType = listEl ? listEl.tagName.toLowerCase() : undefined;
      const listStyles = listEl ? window.getComputedStyle(listEl) : undefined;
      const listStyleType = listStyles?.listStyleType;
      const listStylePosition = listStyles?.listStylePosition as
        | "inside"
        | "outside"
        | undefined;

      const wrapper = document.getElementById("presentation-slides-wrapper");
      let listLevelCount = 0;
      let listIndentPx = 0;
      let currentList = listEl;

      while (currentList && (!wrapper || wrapper.contains(currentList))) {
        listLevelCount += 1;
        const cs = window.getComputedStyle(currentList);
        const paddingLeft = parseFloat(cs.paddingLeft || "0");
        const marginLeft = parseFloat(cs.marginLeft || "0");
        listIndentPx += (isNaN(paddingLeft) ? 0 : paddingLeft) + (isNaN(marginLeft) ? 0 : marginLeft);
        const nextParent = currentList.parentElement;
        currentList = nextParent ? nextParent.closest("ul,ol") : null;
      }

      const listLevel = Math.max(0, listLevelCount - 1);

      let listItemIndex: number | undefined;
      const parent = liEl.parentElement;
      if (parent) {
        const items = Array.from(parent.children).filter(
          (child) => (child as HTMLElement).tagName?.toLowerCase() === "li"
        );
        const index = items.indexOf(liEl);
        listItemIndex = index >= 0 ? index : undefined;
      }

      return {
        isListItem: true,
        listType: listType as "ul" | "ol" | undefined,
        listLevel,
        listStyleType: listStyleType || undefined,
        listStylePosition,
        listIndent: listIndentPx > 0 ? listIndentPx : undefined,
        listItemIndex,
      };
    }

    function parseElementAttributes(el: Element) {
      let tagName = el.tagName.toLowerCase();

      const computedStyles = window.getComputedStyle(el);

      const position = parsePosition(el);

      const shadow = parseShadow(computedStyles);

      const background = parseBackground(computedStyles);

      const border = parseBorder(computedStyles);

      const font = parseFont(computedStyles);

      const lineHeight = parseLineHeight(computedStyles, el);

      const margin = parseMargin(computedStyles);

      const padding = parsePadding(computedStyles);

      const innerText = hasOnlyTextNodes(el)
        ? el.textContent || undefined
        : undefined;

      const zIndex = parseInt(computedStyles.zIndex);
      const zIndexValue = isNaN(zIndex) ? 0 : zIndex;

      const textAlign = computedStyles.textAlign as
        | "left"
        | "center"
        | "right"
        | "justify";
      const objectFit = computedStyles.objectFit as
        | "contain"
        | "cover"
        | "fill"
        | undefined;

      const parsedBackgroundImage = parseBackgroundImage(computedStyles);
      const imageSrc = (el as HTMLImageElement).src || parsedBackgroundImage;

      const borderRadiusValue = parseBorderRadius(computedStyles, el);

      const shape = parseShape(el, borderRadiusValue) as
        | "rectangle"
        | "circle"
        | undefined;

      const textWrap = computedStyles.whiteSpace !== "nowrap";

      const filters = parseFilters(computedStyles);

      const opacity = parseFloat(computedStyles.opacity);
      const elementOpacity = isNaN(opacity) ? undefined : opacity;

      const listInfo = parseListInfo(el);
      const dataEditableId = el.getAttribute("data-editable-id") || undefined;

      return {
        tagName: tagName,
        id: el.id,
        dataEditableId: dataEditableId,
        className:
          el.className && typeof el.className === "string"
            ? el.className
            : el.className
            ? el.className.toString()
            : undefined,
        innerText: innerText,
        opacity: elementOpacity,
        background: background,
        border: border,
        shadow: shadow,
        font: font,
        position: position,
        margin: margin,
        padding: padding,
        zIndex: zIndexValue,
        textAlign: textAlign !== "left" ? textAlign : undefined,
        lineHeight: lineHeight,
        borderRadius: borderRadiusValue,
        imageSrc: imageSrc,
        objectFit: objectFit,
        clip: false,
        overlay: undefined,
        shape: shape,
        connectorType: undefined,
        textWrap: textWrap,
        should_screenshot: false,
        element: undefined,
        filters: filters,
        isListItem: listInfo.isListItem,
        listType: listInfo.listType,
        listLevel: listInfo.listLevel,
        listStyleType: listInfo.listStyleType,
        listStylePosition: listInfo.listStylePosition,
        listIndent: listInfo.listIndent,
        listItemIndex: listInfo.listItemIndex,
      };
    }

    return parseElementAttributes(el);
  });
  return attributes;
}
