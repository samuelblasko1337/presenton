import { ElementHandle } from "puppeteer";

export interface ElementAttributes {
  tagName: string;
  id?: string;
  dataEditableId?: string;
  className?: string;
  innerText?: string;
  opacity?: number;
  background?: {
    color?: string;
    opacity?: number;
  };
  border?: {
    color?: string;
    width?: number;
    opacity?: number;
  };
  shadow?: {
    offset?: [number, number];
    color?: string;
    opacity?: number;
    radius?: number;
    angle?: number;
    spread?: number;
    inset?: boolean;
  },
  font?: {
    name?: string;
    family?: string;
    size?: number;
    weight?: number;
    color?: string;
    italic?: boolean;
  };
  position?: {
    left?: number;
    top?: number;
    width?: number;
    height?: number;
  };
  margin?: {
    top?: number;
    bottom?: number;
    left?: number;
    right?: number;
  };
  padding?: {
    top?: number;
    bottom?: number;
    left?: number;
    right?: number;
  };
  zIndex?: number;
  domPath?: string;
  depth?: number;
  isListItem?: boolean;
  listType?: 'ul' | 'ol';
  listLevel?: number;
  listStyleType?: string;
  listStylePosition?: 'inside' | 'outside';
  listIndent?: number;
  listItemIndex?: number;
  textAlign?: 'left' | 'center' | 'right' | 'justify';
  lineHeight?: number;
  borderRadius?: number[];
  imageSrc?: string;
  objectFit?: 'contain' | 'cover' | 'fill';
  clip?: boolean;
  overlay?: string;
  shape?: 'rectangle' | 'circle';
  connectorType?: string;
  textWrap?: boolean;
  should_screenshot?: boolean;
  element?: ElementHandle<Element>;
  filters?: {
    invert?: number;
    brightness?: number;
    contrast?: number;
    saturate?: number;
    hueRotate?: number;
    blur?: number;
    grayscale?: number;
    sepia?: number;
    opacity?: number;
  };
}

export interface SlideAttributesResult {
  elements: ElementAttributes[];
  backgroundColor?: string;
  speakerNote?: string;
}
