import { ElementAttributes, SlideAttributesResult } from "@/types/element_attibutes";
import {
  PptxSlideModel,
  PptxTextBoxModel,
  PptxAutoShapeBoxModel,
  PptxPictureBoxModel,
  PptxConnectorModel,
  PptxPositionModel,
  PptxFillModel,
  PptxStrokeModel,
  PptxShadowModel,
  PptxFontModel,
  PptxParagraphModel,
  PptxPictureModel,
  PptxObjectFitModel,
  PptxBoxShapeEnum,
  PptxObjectFitEnum,
  PptxAlignment,
  PptxShapeType,
  PptxConnectorType
} from "@/types/pptx_models";

function convertTextAlignToPptxAlignment(textAlign?: string): PptxAlignment | undefined {
  if (!textAlign) return undefined;

  switch (textAlign.toLowerCase()) {
    case 'left':
      return PptxAlignment.LEFT;
    case 'center':
      return PptxAlignment.CENTER;
    case 'right':
      return PptxAlignment.RIGHT;
    case 'justify':
      return PptxAlignment.JUSTIFY;
    default:
      return PptxAlignment.LEFT;
  }
}

const PX_TO_PT = 72 / 96;
const DEFAULT_FONT_SIZE_PX = 16;

function pxToPt(value?: number): number | undefined {
  if (value === undefined || value === null || isNaN(value)) return undefined;
  return value * PX_TO_PT;
}

function pxToPtRounded(value?: number): number {
  const pt = pxToPt(value);
  return pt === undefined ? 0 : Math.round(pt);
}

function pxToPtRoundedOr(value: number | undefined, fallbackPx: number): number {
  const raw = value ?? fallbackPx;
  return Math.round(raw * PX_TO_PT);
}

function convertLineHeightToRelative(lineHeight?: number, fontSize?: number): number | undefined {
  if (!lineHeight) return undefined;

  let calculatedLineHeight = 1.2;
  if (lineHeight < 10) {
    calculatedLineHeight = lineHeight;
  }

  if (fontSize && fontSize > 0) {
    calculatedLineHeight = Math.round((lineHeight / fontSize) * 100) / 100;
  }

  return calculatedLineHeight - 0.3
}

export function mapToPptxFontName(rawName?: string, rawFamily?: string, weight?: number, tagName?: string): string {
  const input = `${rawFamily ?? ""} ${rawName ?? ""}`.toLowerCase();
  const w = weight ?? 400;
  const tag = (tagName ?? "").toLowerCase();
  const raw = rawName ?? "";

  const toEquip = (fontWeight: number) => {
    if (fontWeight >= 600) return "Equip";
    if (fontWeight >= 500) return "Equip Medium";
    return "Equip";
  };
  const toEquipExt = (fontWeight: number) =>
    fontWeight >= 700 ? "Equip Extended ExtraBold" : "Equip Extended Light";

  // Next/font class names (match Equip Extended before Equip)
  if (raw.startsWith("__equipExt")) return toEquipExt(w);
  if (raw.startsWith("__equip")) return toEquip(w);
  if (raw.startsWith("__inter")) return toEquip(w);

  // Equip Extended
  if (input.includes("equip extended") || input.includes("equipext") || input.includes("--font-equip-ext")) {
    return toEquipExt(w);
  }

  // Equip
  if (input.includes("equip") || input.includes("--font-equip")) {
    return toEquip(w);
  }

  // Headline roles (fallback if no explicit font family)
  if (tag === "h1" || tag === "h2") return "Equip Extended ExtraBold";
  if (tag === "h3") return "Equip Extended Light";

  // Fallback to embedded Equip faces only
  return toEquip(w);
}

export function convertElementAttributesToPptxSlides(
  slidesAttributes: SlideAttributesResult[]
): PptxSlideModel[] {
  return slidesAttributes.map((slideAttributes) => {
    const shapes = slideAttributes.elements.map(element => {
      return convertElementToPptxShape(element);
    }).filter(Boolean);

    const slide: PptxSlideModel = {
      shapes: shapes as (PptxTextBoxModel | PptxAutoShapeBoxModel | PptxConnectorModel | PptxPictureBoxModel)[],
      note: slideAttributes.speakerNote
    };

    if (slideAttributes.backgroundColor) {
      slide.background = {
        color: slideAttributes.backgroundColor,
        opacity: 1.0
      };
    }

    return slide;
  });
}

function convertElementToPptxShape(
  element: ElementAttributes
): PptxTextBoxModel | PptxAutoShapeBoxModel | PptxConnectorModel | PptxPictureBoxModel | null {

  if (!element.position) {
    return null;
  }

  if (element.tagName === 'img' || (element.className && typeof element.className === 'string' && element.className.includes('image')) || element.imageSrc) {
    return convertToPictureBox(element);
  }

  if (element.innerText && element.innerText.trim().length > 0) {
    // Use AutoShape model if there's background color and border radius
    if (element.background?.color && element.borderRadius && element.borderRadius.some(radius => radius > 0)) {
      return convertToAutoShapeBox(element);
    }
    return convertToTextBox(element);
  }

  if (element.tagName === 'hr') {
    return convertToConnector(element);
  }

  return convertToAutoShapeBox(element);
}

function convertToTextBox(element: ElementAttributes): PptxTextBoxModel {
  const position: PptxPositionModel = {
    left: pxToPtRounded(element.position?.left),
    top: pxToPtRounded(element.position?.top),
    width: pxToPtRounded(element.position?.width),
    height: pxToPtRounded(element.position?.height)
  };

  const fill: PptxFillModel | undefined = element.background?.color ? {
    color: element.background.color,
    opacity: element.background.opacity ?? 1.0
  } : undefined;

  const font: PptxFontModel | undefined = element.font ? {
    name: mapToPptxFontName(element.font.name, (element.font as any).family, element.font.weight, element.tagName),
    size: pxToPtRoundedOr(element.font.size, DEFAULT_FONT_SIZE_PX),
    font_weight: element.font.weight ?? 400,
    italic: element.font.italic ?? false,
    color: element.font.color ?? "000000"
  } : undefined;

  const paragraph: PptxParagraphModel = {
    spacing: undefined,
    alignment: convertTextAlignToPptxAlignment(element.textAlign),
    font,
    line_height: convertLineHeightToRelative(element.lineHeight, element.font?.size),
    text: element.innerText
  };

  return {
    shape_type: "textbox",
    margin: undefined,
    fill,
    position,
    text_wrap: element.textWrap ?? true,
    paragraphs: [paragraph]
  };
}

function convertToAutoShapeBox(element: ElementAttributes): PptxAutoShapeBoxModel {
  const position: PptxPositionModel = {
    left: pxToPtRounded(element.position?.left),
    top: pxToPtRounded(element.position?.top),
    width: pxToPtRounded(element.position?.width),
    height: pxToPtRounded(element.position?.height)
  };
  const fill: PptxFillModel | undefined = element.background?.color ? {
    color: element.background.color,
    opacity: element.background.opacity ?? 1.0
  } : undefined;

  const stroke: PptxStrokeModel | undefined = element.border?.color ? {
    color: element.border.color,
    thickness: pxToPt(element.border.width ?? 1) ?? 1,
    opacity: element.border.opacity ?? 1.0
  } : undefined;

  const shadow: PptxShadowModel | undefined = element.shadow?.color ? {
    radius: pxToPtRounded(element.shadow.radius ?? 4),
    offset: pxToPtRounded(element.shadow.offset ? Math.sqrt(element.shadow.offset[0] ** 2 + element.shadow.offset[1] ** 2) : 0),
    color: element.shadow.color,
    opacity: element.shadow.opacity ?? 0.5,
    angle: Math.round(element.shadow.angle ?? 0)
  } : undefined;

  const paragraphs: PptxParagraphModel[] | undefined = element.innerText ? [{
    spacing: undefined,
    alignment: convertTextAlignToPptxAlignment(element.textAlign),
    font: element.font ? {
      name: mapToPptxFontName(element.font.name, (element.font as any).family, element.font.weight, element.tagName),
      size: pxToPtRoundedOr(element.font.size, DEFAULT_FONT_SIZE_PX),
      font_weight: element.font.weight ?? 400,
      italic: element.font.italic ?? false,
      color: element.font.color ?? "000000"
    } : undefined,
    line_height: convertLineHeightToRelative(element.lineHeight, element.font?.size),
    text: element.innerText
  }] : undefined;

  const shapeType = element.borderRadius ? PptxShapeType.ROUNDED_RECTANGLE : PptxShapeType.RECTANGLE;

  let borderRadius = undefined;
  for (const eachCornerRadius of element.borderRadius ?? []) {
    if (eachCornerRadius > 0) {
      borderRadius = Math.max(borderRadius ?? 0, pxToPtRounded(eachCornerRadius));
    }
  }

  return {
    shape_type: "autoshape",
    type: shapeType,
    margin: undefined,
    fill,
    stroke,
    shadow,
    position,
    text_wrap: element.textWrap ?? true,
    border_radius: borderRadius || undefined,
    paragraphs
  };
}

function convertToPictureBox(element: ElementAttributes): PptxPictureBoxModel {
  const position: PptxPositionModel = {
    left: pxToPtRounded(element.position?.left),
    top: pxToPtRounded(element.position?.top),
    width: pxToPtRounded(element.position?.width),
    height: pxToPtRounded(element.position?.height)
  };

  const objectFit: PptxObjectFitModel = {
    fit: element.objectFit ? (element.objectFit as PptxObjectFitEnum) : PptxObjectFitEnum.CONTAIN
  };

  const picture: PptxPictureModel = {
    is_network: element.imageSrc ? element.imageSrc.startsWith('http') : false,
    path: element.imageSrc || ''
  };

  return {
    shape_type: "picture",
    position,
    margin: undefined,
    clip: element.clip ?? true,
    invert: element.filters?.invert === 1,
    opacity: element.opacity,
    border_radius: element.borderRadius ? element.borderRadius.map(r => Math.round(r * PX_TO_PT)) : undefined,
    shape: element.shape ? (element.shape as PptxBoxShapeEnum) : PptxBoxShapeEnum.RECTANGLE,
    object_fit: objectFit,
    picture
  };
}

function convertToConnector(element: ElementAttributes): PptxConnectorModel {
  const position: PptxPositionModel = {
    left: pxToPtRounded(element.position?.left),
    top: pxToPtRounded(element.position?.top),
    width: pxToPtRounded(element.position?.width),
    height: pxToPtRounded(element.position?.height)
  };

  return {
    shape_type: "connector",
    type: PptxConnectorType.STRAIGHT,
    position,
    thickness: pxToPt(element.border?.width ?? 0.5) ?? 0.5,
    color: element.border?.color || element.background?.color || '000000',
    opacity: element.border?.opacity ?? 1.0
  };
}
