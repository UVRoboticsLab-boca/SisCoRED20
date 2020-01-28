import { AlignmentType, ISpacingProperties } from "../../../file/paragraph";
import { IIndentAttributesProperties } from "../../../file/paragraph/formatting";
import { UnderlineType } from "../../../file/paragraph/run/underline";
import { ShadingType } from "../../../file/table";
import { Style } from "./style";
export interface IBaseParagraphStyleOptions {
    readonly basedOn?: string;
    readonly next?: string;
    readonly quickFormat?: boolean;
    readonly link?: string;
    readonly semiHidden?: boolean;
    readonly uiPriority?: number;
    readonly unhideWhenUsed?: boolean;
    readonly run?: {
        readonly size?: number;
        readonly bold?: boolean;
        readonly italics?: boolean;
        readonly smallCaps?: boolean;
        readonly allCaps?: boolean;
        readonly strike?: boolean;
        readonly doubleStrike?: boolean;
        readonly subScript?: boolean;
        readonly superScript?: boolean;
        readonly underline?: {
            readonly type?: UnderlineType;
            readonly color?: string;
        };
        readonly color?: string;
        readonly font?: string;
        readonly characterSpacing?: number;
        readonly highlight?: string;
        readonly shadow?: {
            readonly type: ShadingType;
            readonly fill: string;
            readonly color: string;
        };
    };
    readonly paragraph?: {
        readonly alignment?: AlignmentType;
        readonly thematicBreak?: boolean;
        readonly rightTabStop?: number;
        readonly leftTabStop?: number;
        readonly indent?: IIndentAttributesProperties;
        readonly spacing?: ISpacingProperties;
        readonly keepNext?: boolean;
        readonly keepLines?: boolean;
        readonly outlineLevel?: number;
    };
}
export interface IParagraphStyleOptions extends IBaseParagraphStyleOptions {
    readonly id: string;
    readonly name?: string;
}
export declare class ParagraphStyle extends Style {
    private readonly paragraphProperties;
    private readonly runProperties;
    constructor(options: IParagraphStyleOptions);
}
