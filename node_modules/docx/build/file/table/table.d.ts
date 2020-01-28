import { XmlComponent } from "../../file/xml-components";
import { WidthType } from "./table-cell";
import { ITableFloatOptions } from "./table-properties";
import { TableLayoutType } from "./table-properties/table-layout";
import { TableRow } from "./table-row";
export interface ITableOptions {
    readonly rows: TableRow[];
    readonly width?: {
        readonly size: number;
        readonly type?: WidthType;
    };
    readonly columnWidths?: number[];
    readonly margins?: {
        readonly marginUnitType?: WidthType;
        readonly top?: number;
        readonly bottom?: number;
        readonly right?: number;
        readonly left?: number;
    };
    readonly float?: ITableFloatOptions;
    readonly layout?: TableLayoutType;
}
export declare class Table extends XmlComponent {
    private readonly properties;
    constructor({ rows, width, columnWidths, margins: { marginUnitType, top, bottom, right, left }, float, layout, }: ITableOptions);
}
