//@ts-ignore
import { SharepointServiceWorker } from "./Secrets";

export const ProcessPayloadsResponseType = {
  update: 0,
  create: 1,
  delete: 2,
} as const;


/**
 * PowerAutomate service worker payload types
 */
type PowerAutomatePayload =
  | {
      operation: "rest";
      rest: {
        site: string;
        api: string;
        method: string;
        headers?: Record<string, string>;
        body?: string;
      };
      file?: never;
    }
  | {
      operation: "file-create" | "file-update";
      file: {
        site: string;
        folderPath: string; // e.g. "/subFolder/nested" or "/" for root
        fileName: string; // e.g. "myFile.txt"
        fileContentBase64: string; // base64-encoded
      };
      rest?: never;
    };

/**
 * Sharepoint Field Schema definitions
 */

export interface BaseFieldSchema {
  internalName: string;
  nullable?: boolean;
}

export type BasicSharePointFieldType = "string" | "number" | "boolean" | "date";

export type SimpleFieldSchema = BaseFieldSchema & {
  type: BasicSharePointFieldType;
};

export type ChoiceFieldSchema = BaseFieldSchema & {
  type: "choice"; // Specific type literal
  choices?: readonly string[];
};

export const DefaultPersonSelectedProps = [
  "Id",
  "Title",
  "EMail",
  "Name",
  "FirstName",
  "LastName",
  "JobTitle",
  "Department",
  "MobilePhone",
  "WorkPhone",
] as const;
export type TDefaultPersonSelectedProps =
  (typeof DefaultPersonSelectedProps)[number];

export const DefaultLookupSelectedProps = ["Id", "Title"] as const;
export type TDefaultLookupSelectedProps =
  (typeof DefaultLookupSelectedProps)[number];

export const DefaultFileSelectedProps = [
  "Name",
  "ServerRelativeUrl",
  "TimeCreated",
  "TimeLastModified",
  "Length",
  "Title",
  "UniqueId",
  "CheckOutType",
  "ETag",
] as const;
export type TDefaultFileSelectedProps =
  (typeof DefaultFileSelectedProps)[number];

export type PersonFieldSchema<
  SelectProps extends readonly string[] = typeof DefaultPersonSelectedProps
> = BaseFieldSchema & {
  type: "person"; // Renamed from 'user'
  selectProps?: SelectProps;
};

export type MultiPersonFieldSchema<
  SelectProps extends readonly string[] = typeof DefaultPersonSelectedProps
> = BaseFieldSchema & {
  type: "multiperson";
  selectProps?: SelectProps;
};

export type LookupFieldSchema<
  SelectProps extends readonly string[] = typeof DefaultLookupSelectedProps
> = BaseFieldSchema & {
  type: "lookup";
  selectProps?: SelectProps;
};

export type FileFieldSchema<
  SelectProps extends readonly string[] = typeof DefaultFileSelectedProps
> = BaseFieldSchema & {
  type: "file";
  selectProps?: SelectProps;
};

export type MultiChoiceFieldSchema = BaseFieldSchema & {
  type: "multichoice";
  choices?: readonly string[];
};

export type SharePointFieldSchema =
  | SimpleFieldSchema
  | PersonFieldSchema<any>
  | MultiPersonFieldSchema<any>
  | LookupFieldSchema<any>
  | FileFieldSchema<any>
  | ChoiceFieldSchema
  | MultiChoiceFieldSchema;

// --- Restore previously removed type definitions ---
export type SharePointPerson = {
  Id: number;
  Title?: string; // Display Name
  Email?: string;
  Name?: string; // Login Name
  FirstName?: string;
  LastName?: string;
  JobTitle?: string;
  Department?: string;
  MobilePhone?: string;
  WorkPhone?: string;
  Picture?: {
    Description: string;
    Url: string;
  };
  [key: string]: any;
};

export type SharePointLookup = {
  Id: number;
  Title: string; // Display Name of the looked-up item
};

export type SharePointFile = {
  Name: string;
  ServerRelativeUrl: string;
  TimeCreated: string;
  TimeLastModified: string;
  Length: string; // SharePoint returns Length as a string, not a number
  Title?: string | null;
  UniqueId: string;
  CheckOutType: number;
  ETag: string;
  CheckInComment?: string;
  ContentTag?: string;
  CustomizedPageStatus?: number;
  Exists?: boolean;
  ExistsAllowThrowForPolicyFailures?: boolean;
  ExistsWithException?: boolean;
  IrmEnabled?: boolean;
  Level?: number;
  LinkingUri?: string | null;
  LinkingUrl?: string;
  MajorVersion?: number;
  MinorVersion?: number;
  UIVersion?: number;
  UIVersionLabel?: string;
  [key: string]: any;
};

// 1. Map Basic SharePoint Field Types to TypeScript Types
export type MapBasicType<T extends BasicSharePointFieldType> =
  T extends "string"
    ? string
    : T extends "number"
    ? number
    : T extends "boolean"
    ? boolean
    : T extends "date"
    ? string // SharePoint REST API returns dates as strings
    : never;

export type InferMode = "get" | "set";

export type InferFieldType<
  TField extends SharePointFieldSchema,
  mode extends InferMode
> =
  | (TField extends SimpleFieldSchema // If simple field
      ? MapBasicType<TField["type"]>
      : TField extends ChoiceFieldSchema // If Choice field
      ? TField["choices"] extends readonly string[]
        ? TField["choices"][number]
        : string // If choices are not defined, fallback to string
      : TField extends PersonFieldSchema<any> // If Person field
      ? mode extends "get"
        ? SharePointPerson //Return person object
        : string | number //When setting we can use email, userid, display name or Id
      : TField extends LookupFieldSchema<any> // If Lookup field
      ? mode extends "get"
        ? SharePointLookup //Return lookup object
        : number //When setting we can use Id
      : TField extends FileFieldSchema<any> // If File field
      ? mode extends "get"
        ? SharePointFile //Return file object
        : never //File cannot be set directly
      : TField extends MultiChoiceFieldSchema // If MultiChoice field
      ? TField["choices"] extends readonly string[]
        ? TField["choices"][number][]
        : string[] // If choices are not defined, fallback to string[]
      : TField extends MultiPersonFieldSchema<any> // If MultiPerson field
      ? mode extends "get"
        ? SharePointPerson[] //Return array of person objects
        : string[] | number[] //When setting we can use email, userid, display name or Id array
      : never)
  // Apply nullability
  | (TField["nullable"] extends true ? null : never);

export type InferGetterFromSharePointSchema<
  TSchema extends readonly SharePointFieldSchema[]
> = {
  // Map over each field definition (SchemaItem) in the TSchema tuple/array.
  // Each SchemaItem will be one of the concrete types (SimpleFieldSchema, PersonFieldSchema, etc.)
  [SchemaItem in TSchema[number] as SchemaItem["internalName"]]: InferFieldType<
    SchemaItem,
    "get"
  >;
} & {
  Id: number; //Id always present
};

export type InferCreatorFromSharePointSchema<
  TSchema extends readonly SharePointFieldSchema[]
> = // Required fields (not nullable)
  Omit<
    {
      [SchemaItem in TSchema[number] as SchemaItem extends
        | PersonFieldSchema<any>
        | MultiPersonFieldSchema<any>
        | FileFieldSchema<any>
        ? never
        : SchemaItem["nullable"] extends true
        ? never
        : SchemaItem["internalName"]]: InferFieldType<SchemaItem, "set">;
    } & {
      // Optional fields (nullable)
      [SchemaItem in TSchema[number] as SchemaItem extends
        | PersonFieldSchema<any>
        | MultiPersonFieldSchema<any>
        | FileFieldSchema<any>
        ? never
        : SchemaItem["nullable"] extends true
        ? SchemaItem["internalName"]
        : never]?: InferFieldType<SchemaItem, "set">;
    } & {
      // Required person/multiperson fields (not nullable)
      [SchemaItem in TSchema[number] as SchemaItem extends
        | PersonFieldSchema<any>
        | MultiPersonFieldSchema<any>
        ? SchemaItem["nullable"] extends true
          ? never
          : `${SchemaItem["internalName"]}Email`
        : never]: SchemaItem extends MultiPersonFieldSchema<any> ? string[]: string;
    } & {
      // Optional person/multiperson fields (nullable)
      [SchemaItem in TSchema[number] as SchemaItem extends
        | PersonFieldSchema<any>
        | MultiPersonFieldSchema<any>
        ? SchemaItem["nullable"] extends true
          ? `${SchemaItem["internalName"]}Email`
          : never
        : never]?: SchemaItem extends MultiPersonFieldSchema<any> ? string[]: string;
    },
    "Id"
  >;

export type InferSetterFromSharePointSchema<
  TSchema extends readonly SharePointFieldSchema[]
> = Partial<InferCreatorFromSharePointSchema<TSchema>>;

/**
 * Helper function to provide type inference for the schema array literal
 * @param schemaDef sharepoint field schema definitions
 * @returns
 */
export function defineSharePointSchema<T extends SharePointFieldSchema[]>(
  schemaDef: T
): T {
  return schemaDef;
}

// --- Utility type to extract selectProps from schema instance ---
// Used to extract the literal type of selectProps from a schema instance
// If not present, fallback to the default
export type ExtractPersonSelectProps<T> = T extends {
  selectProps: infer Props extends readonly string[];
}
  ? Props
  : typeof DefaultPersonSelectedProps;
export type ExtractLookupSelectProps<T> = T extends {
  selectProps: infer Props extends readonly string[];
}
  ? Props
  : typeof DefaultLookupSelectedProps;

/**
 * Utility type to expand SharePoint field names for $select queries.
 * Handles 'person', 'multiperson', and 'lookup' types by generating 'FieldName/Property' strings.
 * For person/multiperson: always include all default person props unless selectProps is provided.
 * For lookup: always include Id and Title, plus any selectProps if provided.
 */
export type ExpandedSharePointFieldNames<T extends SharePointFieldSchema> =
  T extends PersonFieldSchema<any>
    ? `${T["internalName"]}/${ExtractPersonSelectProps<T>[number]}`
    : T extends MultiPersonFieldSchema<any>
    ? `${T["internalName"]}/${ExtractPersonSelectProps<T>[number]}`
    : T extends LookupFieldSchema<any>
    ? `${T["internalName"]}/${ExtractLookupSelectProps<T>[number]}`
    : T extends FileFieldSchema<any>
    ? `${T["internalName"]}/${ExtractFileSelectProps<T>[number]}`
    : T["internalName"];

// Helper type to extract File selectProps
export type ExtractFileSelectProps<T> = T extends {
  selectProps: infer Props extends readonly string[];
}
  ? Props
  : typeof DefaultFileSelectedProps;

// Special type for File property expansion (for backwards compatibility with direct "File" usage)
export type FilePropertyNames = 
  | "File"
  | `File/${TDefaultFileSelectedProps[number]}`;

// Helper type to append File field to a schema
// Preserves readonly and tuple structure for proper type inference
export type AppendFileToSchema<TSchema extends readonly SharePointFieldSchema[]> = 
  TSchema extends readonly [...infer Rest]
    ? readonly [...Rest, FileFieldSchema<typeof DefaultFileSelectedProps>]
    : readonly [...TSchema, FileFieldSchema<typeof DefaultFileSelectedProps>];

// Helper type to ensure File is visible in autocomplete for document libraries
// This explicitly includes File in the type to help TypeScript's autocomplete
export type DocumentLibraryResult<TSchema extends readonly SharePointFieldSchema[]> = 
  InferGetterFromSharePointSchema<AppendFileToSchema<TSchema>> & {
    File: SharePointFile; // Explicitly include File to ensure it appears in autocomplete
  };

// Operators for string fields
export type ODataFilterStringOperators =
  | "eq"
  | "ne"
  | "gt"
  | "ge"
  | "lt"
  | "le"
  | "startswith"
  | "endswith"
  | "substringof";

// Operators for numeric fields
export type ODataFilterNumericOperators =
  | "eq"
  | "ne"
  | "gt"
  | "ge"
  | "lt"
  | "le";

// Operators for date fields (typically treated as strings by SharePoint OData)
export type ODataFilterDateOperators = "eq" | "ne" | "gt" | "ge" | "lt" | "le";

// Operators for boolean fields
export type ODataFilterBooleanOperators = "eq" | "ne";

// Operators for choice fields (can be treated as strings for 'eq'/'ne')
export type ODataFilterChoiceOperators = "eq" | "ne";

// Operators for Person/Lookup Id fields (numeric)
export type ODataFilterPersonLookupIdOperators =
  | "eq"
  | "ne"
  | "gt"
  | "ge"
  | "lt"
  | "le";

export type ODataFilterAnyOperators =
  | ODataFilterStringOperators
  | ODataFilterNumericOperators
  | ODataFilterDateOperators
  | ODataFilterBooleanOperators
  | ODataFilterChoiceOperators
  | ODataFilterPersonLookupIdOperators;

/**
 * Utility type to get all possible filterable field names (internal and expanded).
 * For person/multiperson: always include all default person props unless selectProps is provided.
 * For lookup: always include Id and Title, plus any selectProps if provided.
 */
export type ODataFilterFieldNames<T extends SharePointFieldSchema> =
  T extends PersonFieldSchema<any>
    ? `${T["internalName"]}/${ExtractPersonSelectProps<T>[number]}`
    : T extends MultiPersonFieldSchema<any>
    ? `${T["internalName"]}/${ExtractPersonSelectProps<T>[number]}`
    : T extends LookupFieldSchema<any>
    ? `${T["internalName"]}/${ExtractLookupSelectProps<T>[number]}`
    : T["internalName"];

// Helper type: for a given prop, is it the Id prop?
type IsIdProp<P extends string> = P extends "Id" ? true : false;

// Restore GetFieldDetails type for use in FilterFieldDetails
export type GetFieldDetails<TSchema extends readonly SharePointFieldSchema[]> =
  {
    [K in TSchema[number] as K["internalName"]]: K;
  } & { Id: null };

/**
 * FilterFieldDetails: maps all possible filterable field names to their value/operator types.
 * For expanded person/lookup fields:
 *   - 'Id' is always number with numeric operators
 *   - all others are string with string operators
 */
export type FilterFieldDetails<
  TSchema extends readonly SharePointFieldSchema[]
> = {
  [FilterName in
    | ODataFilterFieldNames<TSchema[number]>
    | "Id"]: FilterName extends keyof GetFieldDetails<TSchema> // Internal names
    ? FilterName extends "Id"
      ? { valueType: number; operators: ODataFilterNumericOperators }
      : GetFieldDetails<TSchema>[FilterName] extends SimpleFieldSchema & {
          type: "string";
        }
      ? { valueType: string; operators: ODataFilterStringOperators }
      : GetFieldDetails<TSchema>[FilterName] extends SimpleFieldSchema & {
          type: "number";
        }
      ? { valueType: number; operators: ODataFilterNumericOperators }
      : GetFieldDetails<TSchema>[FilterName] extends SimpleFieldSchema & {
          type: "boolean";
        }
      ? { valueType: boolean; operators: ODataFilterBooleanOperators }
      : GetFieldDetails<TSchema>[FilterName] extends SimpleFieldSchema & {
          type: "date";
        }
      ? { valueType: string; operators: ODataFilterDateOperators }
      : GetFieldDetails<TSchema>[FilterName] extends ChoiceFieldSchema
      ? {
          valueType: GetFieldDetails<TSchema>[FilterName]["choices"] extends readonly string[]
            ? GetFieldDetails<TSchema>[FilterName]["choices"][number]
            : string;
          operators: ODataFilterChoiceOperators;
        }
      : GetFieldDetails<TSchema>[FilterName] extends MultiChoiceFieldSchema
      ? {
          valueType: GetFieldDetails<TSchema>[FilterName]["choices"] extends readonly string[]
            ? GetFieldDetails<TSchema>[FilterName]["choices"][number]
            : string;
          operators: ODataFilterChoiceOperators;
        }
      : GetFieldDetails<TSchema>[FilterName] extends PersonFieldSchema<any>
      ? { valueType: number; operators: ODataFilterNumericOperators }
      : GetFieldDetails<TSchema>[FilterName] extends LookupFieldSchema<any>
      ? { valueType: number; operators: ODataFilterNumericOperators }
      : GetFieldDetails<TSchema>[FilterName] extends MultiPersonFieldSchema<any>
      ? { valueType: number; operators: ODataFilterNumericOperators }
      : never
    : // Expanded names
    FilterName extends `${infer FieldName}/${infer Prop}`
    ? FieldName extends keyof GetFieldDetails<TSchema>
      ? GetFieldDetails<TSchema>[FieldName] extends PersonFieldSchema<any>
        ? IsIdProp<Prop> extends true
          ? { valueType: number; operators: ODataFilterNumericOperators }
          : Prop extends TDefaultPersonSelectedProps
          ? { valueType: string; operators: ODataFilterStringOperators }
          : { valueType: any; operators: ODataFilterAnyOperators }
        : GetFieldDetails<TSchema>[FieldName] extends LookupFieldSchema<any>
        ? IsIdProp<Prop> extends true
          ? { valueType: number; operators: ODataFilterNumericOperators }
          : Prop extends TDefaultLookupSelectedProps
          ? { valueType: string; operators: ODataFilterStringOperators }
          : { valueType: any; operators: ODataFilterAnyOperators }
        : GetFieldDetails<TSchema>[FieldName] extends MultiPersonFieldSchema<any>
        ? IsIdProp<Prop> extends true
          ? { valueType: number; operators: ODataFilterNumericOperators }
          : Prop extends TDefaultPersonSelectedProps
          ? { valueType: string; operators: ODataFilterStringOperators }
          : { valueType: any; operators: ODataFilterAnyOperators }
        : never
      : never
    : never;
};

export type ODataFilter<T extends readonly SharePointFieldSchema[]> =
  | ODataLogicalFilter<T>
  | ODataSingleFieldFilter<T>; // Renamed from FieldFilter to clarify it's one field's filter

export type ODataSingleFieldFilter<T extends readonly SharePointFieldSchema[]> =
  {
    [K in keyof FilterFieldDetails<T>]: {
      field: K;
      operator: FilterFieldDetails<T>[K] extends { operators: infer O }
        ? O
        : never;
      value:
        | (FilterFieldDetails<T>[K] extends { valueType: infer V } ? V : never)
        // Special handling for nullability
        | (K extends `${infer N}/${infer P}`
            ? never // Expanded fields are generally not nullable themselves, but their parent might be (handled below)
            : T[number]["nullable"] extends true // Check if any schema has nullable true (this is a bit broad, ideally check specific field)
            ? null
            : never);
    };
  }[keyof FilterFieldDetails<T>];

export type ODataLogicalFilter<T extends readonly SharePointFieldSchema[]> =
  | {
      and: ODataFilter<T>[];
    }
  | {
      or: ODataFilter<T>[];
    };

/**
 * Defines the structure for an OData query against a SharePoint list,
 * with type-safe $select for expanded properties.
 */
export type ODataQuery<T extends readonly SharePointFieldSchema[]> = {
  /**
   * Specifies which properties to return.
   * Example: "$select=id,name,RiskRaiser/Email,File/Name"
   */
  $select?: Array<ExpandedSharePointFieldNames<T[number]> | "Id" | FilePropertyNames> | "*"; // Id is always selectable | "*" selects all fields

  /**
   * Specifies which related objects to expand.
   * Example: "$expand=RiskRaiser,ClosedBy,File"
   * Only the internalName of the person/lookup field is used here, or "File" for document libraries.
   */
  $expand?: Array<
    | {
        [K in keyof T]: T[K] extends { type: "person" | "multiperson" | "lookup" }
          ? T[K]["internalName"]
          : never;
      }[number]
    | "File"
  >;

  /**
   * Specifies the number of items to skip.
   * Example: "$skip=10"
   */
  $skip?: number;

  /**
   * Specifies filtering criteria.
   * Example: [
   *  and: [
   *     { field: "Id", operator: "gt", value: 100 },
   *     { field: "Description", operator: "substringof", value: "critical" }
   *  ]
   * ]
   */
  $filter?: ODataFilter<T>;

  /**
   * Specifies the order of returned items.
   * Example: "$orderby=Title asc, Created desc"
   */
  $orderby?: Array<{
    name: ExpandedSharePointFieldNames<T[number]> | "Id" | FilePropertyNames;
    direction: "asc" | "desc";
  }>;

  /**
   * Specifies the maximum number of items to return.
   * Example: "$top=50"
   */
  $top?: number;
  /**
   * Specifies whether to include item count.
   * Example: "$inlinecount=allpages"
   */
  $inlinecount?: "allpages" | "none";
  /**
   * Specifies what data to include (e.g., 'properties', 'metadata').
   * Rarely used with basic SharePoint REST.
   */
  $format?: string;
  /**
   * Specifies specific list item versions.
   * Example: "$select=Title&$select=Version,Created"
   */
  $selectversion?: boolean; // SharePoint specific
};

/**
 * Utility type to extract only the columns specified in the ODataQuery's $select property.
 * If $select is "*", all columns from the schema are included.
 * If $select is omitted, all columns from the schema are included.
 * Otherwise, only the selected columns are included.
 */
export type ODataQueryResult<
  TSchema extends readonly SharePointFieldSchema[],
  TQuery extends
    | { $select?: ReadonlyArray<string> | "*" | undefined }
    | undefined,
  FullResult = InferGetterFromSharePointSchema<TSchema>,
  Selections = TQuery extends { $select: readonly (infer S)[] } ? S : never
> = TQuery extends undefined
  ? FullResult
  : TQuery["$select"] extends "*" | undefined
  ? FullResult
  : Selections extends string
  ? Pick<FullResult, Extract<Selections, keyof FullResult>> & {
      [K in (Selections extends `${infer P}/${string}` ? P : never) &
        keyof FullResult]: FullResult[K] extends object
        ? Pick<
            FullResult[K],
            (Selections extends `${K & string}/${infer Prop}` ? Prop : never) &
              keyof FullResult[K]
          >
        : FullResult[K];
    }
  : never;

/**
 * Converts an ODataQuery object into an OData query string for SharePoint REST API.
 * @param query The ODataQuery object
 * @returns The OData query string (e.g., "?$top=10&$orderby=Title asc")
 */
export function odataQueryToString<T extends readonly SharePointFieldSchema[]>(
  query?: ODataQuery<T>,
  schema?: readonly SharePointFieldSchema[]
): string {
  const params: string[] = [];

  if (!query || query.$select === "*" || query.$select === undefined) {
    //Obtain field names from schema
    if (schema) {
      const fieldNames = schema.flatMap((f) => {
        if (f.type === "person" || f.type === "multiperson") {
          const props =
            f.selectProps && f.selectProps.length > 0
              ? f.selectProps
              : DefaultPersonSelectedProps;
          return props.map((p) => `${f.internalName}/${p}`);
        }
        if (f.type === "lookup" && f.selectProps) {
          return f.selectProps.map((p) => `${f.internalName}/${p}`);
        }
        if (f.type === "file") {
          const props =
            f.selectProps && f.selectProps.length > 0
              ? f.selectProps
              : DefaultFileSelectedProps;
          // ListItemAllFields is not a File property - it's a list item property
          // Filter it out when selecting File properties
          const fileProps = props.filter((p) => p !== "ListItemAllFields");
          return fileProps.map((p) => `${f.internalName}/${p}`);
        }
        return f.internalName;
      });
      params.push(
        `$select=${[...new Set(fieldNames)].map(encodeURIComponent).join(",")}`
      );
    }
  } else {
    params.push(`$select=${query.$select.map(encodeURIComponent).join(",")}`);
  }

  if (!query || query.$select == "*" || query.$select === undefined) {
    //Auto-expand file, person and lookup fields if specified in schema
    if (schema) {
      params.push(
        `$expand=${schema
          .filter((e) => ["person", "multiperson", "lookup", "file"].includes(e.type))
          .map((e) => encodeURIComponent(e.internalName))
          .join(",")}`
      );
    }
  } else {
    if (query.$expand && query.$expand.length > 0) {
      params.push(`$expand=${query.$expand.map(encodeURIComponent).join(",")}`);
    }
  }

  if (!!query) {
    if (query.$top) {
      params.push(`$top=${query.$top}`);
    }

    if (query.$skip) {
      params.push(`$skip=${query.$skip}`);
    }

    if (
      query.$orderby &&
      Array.isArray(query.$orderby) &&
      query.$orderby.length > 0
    ) {
      const orderbyStr = query.$orderby
        .map((o) => `${encodeURIComponent(o.name)} ${o.direction}`)
        .join(", ");
      params.push(`$orderby=${orderbyStr}`);
    }

    if (query.$filter) {
      const filterStr = odataFilterToString(query.$filter);
      if (filterStr) {
        params.push(`$filter=${encodeURIComponent(filterStr)}`);
      }
    }

    if (query.$inlinecount) {
      params.push(`$inlinecount=${query.$inlinecount}`);
    }

    if (query.$format) {
      params.push(`$format=${encodeURIComponent(query.$format)}`);
    }

    if (query.$selectversion) {
      params.push(`$selectversion=true`);
    }
  }

  return params.length > 0 ? "?" + params.join("&") : "";
}

/**
 * Converts an ODataFilter object into an OData $filter string.
 * Handles basic and/or/not logic and simple field comparisons.
 * @param filter The ODataFilter object
 * @returns The OData $filter string
 * @see https://www.odata.org/documentation/odata-version-2-0/uri-conventions/#FilterSystemQueryOption
 */
function odataFilterToString<T extends readonly SharePointFieldSchema[]>(
  filter: ODataFilter<T>
): string {
  if (!filter) return "";

  if ("and" in filter && Array.isArray(filter.and)) {
    return filter.and.map(odataFilterToString).filter(Boolean).join(" and ");
  }
  if ("or" in filter && Array.isArray(filter.or)) {
    return filter.or.map(odataFilterToString).filter(Boolean).join(" or ");
  }

  if ("field" in filter) {
    const field = filter.field;
    const op = filter.operator;
    const value = filter.value;

    switch (op) {
      case "eq":
      case "ne":
      case "gt":
      case "ge":
      case "lt":
      case "le":
        return `${field} ${op} ${odataValueToString(value)}`;
      case "in":
        if (Array.isArray(value) && value.length > 0) {
          return (
            "(" +
            value
              .map((v) => `${field} eq ${odataValueToString(v)}`)
              .join(" or ") +
            ")"
          );
        }
        return "";
      case "substringof":
        return `substringof(${odataValueToString(value)}, ${field})`;
      case "startswith":
        return `startswith(${field}, ${odataValueToString(value)})`;
      case "endswith":
        return `endswith(${field}, ${odataValueToString(value)})`;
      case "length":
        return `length(${field})`;

      default:
        return "";
    }
  }
}

/**
 * Converts a value to an OData-compatible string.
 * Handles string escaping, dates, booleans, and numbers.
 * @param value The value to convert
 * @returns The OData string representation
 */
function odataValueToString(value: any): string {
  if (typeof value === "string") {
    // Escape single quotes by doubling them
    return `'${value.replace(/'/g, "''")}'`;
  }
  if (typeof value === "number" || typeof value === "boolean") {
    return String(value);
  }
  if (value instanceof Date) {
    return value.toISOString();
  }
  if (value === null) {
    return "null";
  }
  return String(value);
}

/**
 * Parses the body of the response from the batch request and returns an array of objects with the status code, status text, and JSON data
 * @param body - The body of the response from the batch request
 * @returns An array of objects with the status code, status text, and JSON data
 */
function parseODataBatchResponse(body: string): {statusCode: number, statusText: string, json: any}[] {
  // Detect the boundary
  const boundaryMatch = body.match(/--batchresponse_[\w-]+/);
  if (!boundaryMatch) throw new Error("No batch boundary found");

  const boundary = boundaryMatch[0];
  const parts = body.split(boundary).filter(p => p.trim().length > 0 && p.trim() !== '--');

  return parts.map(part => {
    const statusMatch = part.match(/HTTP\/\d\.\d\s+(\d+)\s+(.+)/);
    const statusCode = statusMatch ? parseInt(statusMatch[1], 10) : null;
    const statusText = statusMatch ? statusMatch[2].trim() : null;

    // Try to extract any JSON block
    const jsonMatch = part.match(/\{[\s\S]*\}/);
    let json = null;
    if (jsonMatch) {
      try {
        json = JSON.parse(jsonMatch[0]);
      } catch {
        // sometimes there is no JSON body
        json = null;
      }
    }

    return { statusCode, statusText, json };
  });
}

/*
 * Create a new GUID
 * @returns {string} A newly generated UUID
 */
function generateUUID(): string {
  return "10000000-1000-4000-8000-100000000000".replace(/[018]/g, (c) =>
    (
      parseInt(c, 10) ^
      (crypto.getRandomValues(new Uint8Array(1))[0] &
        (15 >> (parseInt(c, 10) / 4)))
    ).toString(16)
  );
}

/*
 * Split an array into chunks of a specified size
 * @param {T[]} array - the array to split into chunks
 * @param {number} chunkSize - the chunk size desired
 * @returns {T[][]} - Array of chunks
 */
function splitArrayIntoChunks<T>(array: T[], chunkSize: number = 250): T[][] {
  const result: T[][] = [];
  for (let i = 0; i < array.length; i += chunkSize) {
    result.push(array.slice(i, i + chunkSize));
  }
  return result;
}

// Define an interface for the items to be changed in a batch update
export interface ISharePointListChange<
  TSchema extends readonly SharePointFieldSchema[]
> {
  id: number;
  data: InferSetterFromSharePointSchema<TSchema>;
}

/**
 * Represents a conflict detected when committing changes to SharePoint.
 * A conflict occurs when:
 * - The user changed a field (oldValue !== newValue)
 * - The database also changed that field (dbValue !== oldValue)
 * - The user's change differs from the database change (newValue !== dbValue)
 */
export type IConflict<T> = {
  itemId: number;
  fieldName: string;
  oldValue: T; // Value before user changed it (from cache)
  newValue: T; // Value after user changed it (from payload)
  conflictValue: T; // New value in the database that caused the conflict
  resolve(value: T): void; // Function to resolve the conflict by setting the value to the specified value
};

export class SharePointSite {
  private site: string;
  private personCache: Map<string, SharePointPerson> = new Map();
  private personPromiseCache: Map<string, Promise<SharePointPerson>> = new Map();

  /**
   * Get a person from their email address.
   * @param email The email of the person to get.
   * @returns The SharePointPerson object.
   */
  async getPersonFromEmail(email: string): Promise<SharePointPerson> {
    // Check cache first for resolved responses
    if (this.personCache.has(email)) {
      return this.personCache.get(email)!;
    }

    // Check if there's an in-flight promise for this email
    if (this.personPromiseCache.has(email)) {
      return this.personPromiseCache.get(email)!;
    }

    // Create a new promise for this request
    const promise = (async () => {
      try {
        let response = this._serviceWorkerFetchWrapper(
          `/_api/web/ensureuser`,
          {
            Accept: "application/json;odata=verbose",
            "Content-Type": "application/json;odata=verbose",
          },
          "POST",
          {
            logonName: email,
          }
        );
        let jsonEnsured = await (await response).json() as any;
        
        // TODO: Load the image, this approach didn't seem to work... I assume I need to make a custom PA Flow to obtain the image from O365 data.
        // const responseImage = await this._serviceWorkerFetchWrapper(
        //   `/_layouts/15/userphoto.aspx?size=L&username=${jsonEnsured.d.Email || encodeURIComponent(jsonEnsured.d.LoginName)}`,
        //   {
        //     Accept: "image/jpeg",
        //   },
        //   "GET"
        // );
        // const image = await responseImage.blob();
        // const blob = new Blob([image], {type: "image/jpeg"});
        // const imageUrl = URL.createObjectURL(blob);
        // debugger
        const imageUrl = "";
        // Normalize email field: SharePoint returns EMail but we use Email
        const person = {
          ...jsonEnsured.d,
          Email: jsonEnsured.d.EMail || jsonEnsured.d.Email || email, // Map EMail to Email, fallback to provided email
          Picture: {
            Description: jsonEnsured.d.Title,
            Url: imageUrl,
          },
        } as SharePointPerson;
        
        // Cache the resolved result
        this.personCache.set(email, person);
        
        return person;
      } finally {
        // Remove from in-flight cache once resolved (whether success or failure)
        this.personPromiseCache.delete(email);
      }
    })();

    // Store the in-flight promise
    this.personPromiseCache.set(email, promise);
    
    return promise;
  }

  constructor(site: string) {
    this.site = site;
  }

  /**
   * Internal helper to make calls to the service worker fetch.
   * This method encapsulates the serviceWorkerFetch logic.
   * @param {string} api - The SharePoint REST API endpoint (relative path, e.g., '/_api/...').
   * @param {Record<string, string>} [headers={}] - Request headers.
   * @param {string} [method="GET"] - HTTP method (GET, POST, etc.).
   * @param {any} [body] - Request body.
   * @param {string} [bodyType="object"] - The type of body the function has been given. 
   * "object" means the body is an object that will be JSON.stringify'd by _serviceWorkerFetchWrapper.
   * "text" means the body is a string that will be sent as is by _serviceWorkerFetchWrapper.
   * @returns {Promise<Response>} - The response from the service worker fetch.
   */
  private async _serviceWorkerFetchWrapper(
    api: string,
    headers: Record<string, string> = {},
    method: string = "GET",
    body?: any,
    bodyType: "object" | "text" = "object"
  ): Promise<Response> {
    // Build the REST operation payload
    const payload: PowerAutomatePayload = {
      operation: "rest",
      rest: {
        site: this.site,
        api,
        method,
        headers: Object.keys(headers).length > 0 ? headers : undefined,
        body: body ? bodyType === "object" ? JSON.stringify(body) : body : undefined,
      },
    };
    const res = await fetch(SharepointServiceWorker, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify(payload),
    });

    if (!res.ok) {
      throw {
        message: `Error fetching from SharePoint: ${res.status} ${res.statusText} for '${api}'`,
        data: await res.json(),
      };
    }
    return res;
  }
}

export class SharePointList<TSchema extends readonly SharePointFieldSchema[]> {
  site: string; // Base SharePoint site URL, also used for the service worker's 'site' parameter
  list: string;
  schema: TSchema;
  private itemType?: Promise<string>; // Promise for in-flight requests to prevent race conditions
  indexes: {
    [K in TSchema[number]["internalName"]]?: Record<string, number[]>;
  } = {};
  private userIdCache: Map<string, number> = new Map();
  protected itemCache: Map<number, InferGetterFromSharePointSchema<TSchema>> =
    new Map();
  private formDigestCache?: { value: string; expiresAt: number; promise?: Promise<string> };

  /**
   * @constructor
   * @param {string} site - The base URL of the SharePoint site (e.g., "https://severntrent.sharepoint.com/sites/OPERATIONALRISKMANAGEMENT/STORM")
   * @param {string} listName - The name of the SharePoint list (e.g., "STORM Missing Data")
   * @example new SharePointList("https://severntrent.sharepoint.com/sites/Asset/Planning/Risks/Ops", "Risks")
   */
  constructor(site: string, listName: string, schema: TSchema) {
    this.site = site;
    this.list = listName;
    this.schema = schema;
  }

  /**
   *
   * @param indexName
   * @param fieldName
   */
  public async createIndex(
    fieldName: TSchema[number]["internalName"]
  ): Promise<void> {
    this.indexes[fieldName] ||= {};
    let items = await this.getItems();
    items.forEach((item) => {
      let key = item[fieldName];
      const index = this.indexes[fieldName] as Record<string, number[]>;
      const keyStr = String(key);
      index[keyStr] ||= [];
      index[keyStr].push(item.Id);
    });
  }

  /**
   * Internal helper to make calls to the service worker fetch for REST operations.
   * This method encapsulates the serviceWorkerFetch logic for standard REST API calls.
   * @param {string} api - The SharePoint REST API endpoint (relative path, e.g., '/_api/...').
   * @param {Record<string, string>} [headers={}] - Request headers.
   * @param {string} [method="GET"] - HTTP method (GET, POST, etc.).
   * @param {any} [body] - Request body (will be JSON stringified).
   * @returns {Promise<Response>} - The response from the service worker fetch.
   */
  protected async _serviceWorkerFetchWrapper(
    api: string,
    headers: Record<string, string> = {},
    method: string = "GET",
    body?: any,
    bodyType: "object" | "text" = "object"
  ): Promise<Response> {
    // Build the REST operation payload
    const payload: PowerAutomatePayload = {
      operation: "rest",
      rest: {
        site: this.site,
        api,
        method,
        headers: Object.keys(headers).length > 0 ? headers : undefined,
        body: body ? bodyType === "object" ? JSON.stringify(body) : body : undefined,
      },
    };

    const res = await fetch(SharepointServiceWorker, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify(payload),
    });

    if (!res.ok) {
      throw {
        message: `Error fetching from SharePoint: ${res.status} ${res.statusText} for '${api}'`,
        data: await res.json(),
      };
    }
    return res;
  }

  /**
   * Get the form digest value from SharePoint contextinfo endpoint.
   * This is required for batch operations that modify data.
   * The digest is cached based on the FormDigestTimeoutSeconds value from SharePoint
   * to avoid excessive API calls.
   * If multiple calls happen concurrently, they will all await the same Promise.
   * @returns Promise resolving to the digest value string
   */
  protected async _getFormDigest(): Promise<string> {
    // Check if we have a cached digest that's still valid
    const now = Date.now();
    if (this.formDigestCache && this.formDigestCache.expiresAt > now) {
      return this.formDigestCache.value;
    }

    // If there's already a Promise in flight, await it
    if (this.formDigestCache?.promise) {
      return await this.formDigestCache.promise;
    }

    // Create a new Promise, cache it, and await it
    const promise = (async () => {
      try {
        // Fetch new digest from SharePoint
        const api = `/_api/contextinfo`;
        const resp = await this._serviceWorkerFetchWrapper(
          api,
          {
            "Accept": "application/json;odata=verbose",
            "Content-Type": "application/json;odata=verbose",
          },
          "POST"
        );
        const json = await resp.json();
        const contextInfo = json.d.GetContextWebInformation as { FormDigestValue: string; FormDigestTimeoutSeconds: number, LibraryVersion: string, SiteFullUrl: string, WebFullUrl: string };
        const expiresAt = Date.now() + (contextInfo.FormDigestTimeoutSeconds - 60) * 1000;
        this.formDigestCache = {
          value: contextInfo.FormDigestValue,
          expiresAt,
        };
        return this.formDigestCache.value;
      } finally {
        // Clear the Promise after it resolves so a new one can be created when needed
        if (this.formDigestCache) {
          this.formDigestCache.promise = undefined;
        }
      }
    })();

    // Store the promise in the cache (create cache object if it doesn't exist)
    if (!this.formDigestCache) {
      this.formDigestCache = { value: "", expiresAt: 0 };
    }
    this.formDigestCache.promise = promise;

    return await promise;
  }

  /**
   * Internal helper to make file operations (create/update) via the service worker.
   * This method handles file uploads using the file-create or file-update operation.
   * @param {File | Blob} file - The file to upload
   * @param {string} folderPath - The folder path relative to the document library (e.g., "/subFolder" or "/" for root)
   * @param {string} fileName - The name of the file
   * @param {"file-create" | "file-update"} operation - The operation type (create or update)
   * @returns {Promise<Response>} - The response from the service worker fetch
   */
  protected async _serviceWorkerFetchWrapperFile(
    file: File | Blob,
    folderPath: string,
    fileName: string,
    operation: "file-create" | "file-update"
  ): Promise<Response> {
    // Convert file to base64
    const arrayBuffer = await file.arrayBuffer();
    const uint8Array = new Uint8Array(arrayBuffer);
    // Convert to binary string efficiently by building in chunks
    const chunkSize = 8192;
    const chunks: string[] = [];
    for (let i = 0; i < uint8Array.length; i += chunkSize) {
      const chunk = uint8Array.slice(i, i + chunkSize);
      // Use String.fromCharCode.apply for better performance on chunks
      chunks.push(String.fromCharCode.apply(null, Array.from(chunk)));
    }
    const data = chunks.join('');
    const fileContentBase64 = btoa(data);

    // Build the file operation payload
    const payload: PowerAutomatePayload = {
      operation,
      file: {
        site: this.site,
        folderPath: `/${this.list}` + (folderPath.startsWith('/') ? "" : "/") + folderPath,
        fileName,
        fileContentBase64,
      },
    };

    const res = await fetch(SharepointServiceWorker, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify(payload),
    });

    if (!res.ok) {
      throw {
        message: `Error with file operation: ${res.status} ${res.statusText} for '${operation}' operation on '${fileName}'`,
        data: await res.json(),
      };
    }
    return res;
  }

  /**
   * Retrieves the sharepoint ID for a field
   * @param userText - The email or login name or member ID of the user to obtain the sharepoint ID for
   * @returns
   */
  public async getOrCreateUserID(userText: string): Promise<number> {
    if (this.userIdCache.has(userText)) {
      return this.userIdCache.get(userText)!;
    }
    const user = await this._serviceWorkerFetchWrapper(
      `/_api/web/ensureuser`,
      {
        Accept: "application/json;odata=verbose",
        "Content-Type": "application/json;odata=verbose",
      },
      "POST",
      { logonName: userText }
    );
    const userJSON = await user.json();
    const id = userJSON.d.Id;
    this.userIdCache.set(userText, id);
    return id;
  }

  /**
   * Transforms multichoice fields from SharePoint format to array format.
   * SharePoint returns: { __metadata: {...}, results: [...] }
   * We transform to: [...]
   * @param item The raw item from SharePoint
   * @returns The item with multichoice fields transformed
   */
  protected _transformGetterDataFromSharePoint(
    item: any
  ): InferGetterFromSharePointSchema<TSchema> {
    const transformed = { ...item };
    
    // Transform multichoice fields
    for (const field of this.schema) {
      if (field.type === "multichoice" && transformed[field.internalName]) {
        const fieldValue = transformed[field.internalName];
        // Check if it's in SharePoint format (has __metadata and results)
        if (
          fieldValue &&
          typeof fieldValue === "object" &&
          "__metadata" in fieldValue &&
          "results" in fieldValue &&
          Array.isArray(fieldValue.results)
        ) {
          transformed[field.internalName] = fieldValue.results;
        }
        // If it's already an array or null/undefined, leave it as is
      } else if (field.type === "person" && transformed[field.internalName]) {
        // Normalize person field: SharePoint returns EMail but we use Email
        const personValue = transformed[field.internalName];
        if (personValue && typeof personValue === "object") {
          transformed[field.internalName] = {
            ...personValue,
            Email: personValue.EMail || personValue.Email || null,
          };
        }
      } else if (field.type === "multiperson" && transformed[field.internalName]) {
        // Normalize multiperson field: SharePoint returns EMail but we use Email
        const personArray = transformed[field.internalName];
        if (Array.isArray(personArray)) {
          transformed[field.internalName] = personArray.map((person: any) => {
            if (person && typeof person === "object") {
              return {
                ...person,
                Email: person.EMail || person.Email || null,
              };
            }
            return person;
          });
        }
      }
    }
    
    return transformed as InferGetterFromSharePointSchema<TSchema>;
  }

  /**
   * Formats a date value to SharePoint format: "YYYY-MM-DD hh:mm:ss"
   * Handles ISO format strings, Date objects, and already-formatted SharePoint dates.
   * @param value The date value to format (can be ISO string, Date object, or SharePoint format string)
   * @returns The date in SharePoint format "YYYY-MM-DD hh:mm:ss", or null if value is null/undefined
   */
  private _formatDateForSharePoint(value: Date | string | null): string | null {
    if (!value) return null;

    // Already in SharePoint format? Return as is.
    let date: Date;
    try {
      date = new Date(value);
    } catch (error) {
      return null;
    }
    
    // Convert date to ISO string, then format to "YYYY-MM-DD hh:mm:ss" by find/replace.
    // This handles cases where the date is in UTC (which .toISOString() yields).
    // The ISO string format is "YYYY-MM-DDTHH:MM:SS.sssZ"
    try {
      const iso = date.toISOString();

      // Replace "T" with space, remove ms and trailing "Z"
      return iso.replace("T", " ").replace(/\.\d{3}Z$/, "");
    } catch (error) {
      console.log("Invalid date format: " + value);
      throw error;
    }
  }

  /**
   * Transforms multichoice fields from array format to SharePoint format.
   * We have: [...]
   * SharePoint expects: { __metadata: { type: 'Collection(Edm.String)' }, results: [...] }
   * @param data The input data object
   * @returns A new object with multichoice fields in SharePoint format
   */
  private _transformMultichoiceFieldsForSharePoint(
    data: Record<string, any>
  ): Record<string, any> {
    const transformed = { ...data };
    
    // Transform multichoice fields
    for (const field of this.schema) {
      if (field.type === "multichoice" && transformed[field.internalName] !== undefined) {
        const fieldValue = transformed[field.internalName];
        // If it's already in SharePoint format, leave it as is
        if (
          fieldValue &&
          typeof fieldValue === "object" &&
          "__metadata" in fieldValue &&
          "results" in fieldValue
        ) {
          // Already in correct format, keep it
          continue;
        }
        // If it's an array, wrap it in SharePoint format
        if (Array.isArray(fieldValue)) {
          transformed[field.internalName] = {
            __metadata: { type: "Collection(Edm.String)" },
            results: fieldValue,
          };
        } else if (fieldValue === null || fieldValue === undefined) {
          // Null/undefined values are fine, leave them as is
          continue;
        }
      }
    }
    
    return transformed;
  }

  /**
   * Transforms setter/creator data by converting all `${internalName}Email` fields for person fields
   * into the correct SharePoint ID property, using getOrCreateUserID. Removes the Email property.
   * Also transforms date fields from ISO format to SharePoint format (YYYY-MM-DD hh:mm:ss)
   * and multichoice fields to SharePoint format.
   * @param data The input data object
   * @returns A new object with SharePoint-ready fields
   */
  protected async _transformSetterDataForSharePoint(
    data:
      | InferSetterFromSharePointSchema<TSchema>
      | InferCreatorFromSharePointSchema<TSchema>
  ): Promise<Record<string, any>> {
    const transformed: Record<string, any> = { ...data };
    const personLookups: Array<Promise<{ idKey: string; userId: number }>> = [];
    const multiPersonLookups: Array<Promise<{ idKey: string; userIds: number[] }>> = [];

    for (const field of this.schema) {
      if (field.type === "person") {
        const emailKey = `${field.internalName}Email`;
        if (emailKey in transformed) {
          if (transformed[emailKey]){
            const idKey = `${field.internalName}Id`;
            // Start the lookup, but don't await yet
            personLookups.push(
              this.getOrCreateUserID(transformed[emailKey]).then((userId) => ({
                idKey,
                userId,
              }))
            );
          }
          //Always delete the email key
          delete transformed[emailKey];
        }
      } else if (field.type === "multiperson") {
        const emailKey = `${field.internalName}Email`;
        if (emailKey in transformed){
          if (transformed[emailKey]) {
            const idKey = `${field.internalName}Id`;
            const emails = transformed[emailKey] as string[] | string;
            const list = Array.isArray(emails) ? emails : [emails];
            multiPersonLookups.push(
              Promise.all(list.map((e) => this.getOrCreateUserID(e))).then((userIds) => ({
                idKey,
                userIds,
              }))
            );
          }
          //Always delete the email key
          delete transformed[emailKey];
        }
      } else if (field.type === "date") {
        // Transform date fields from ISO format to SharePoint format
        if (field.internalName in transformed && transformed[field.internalName] !== undefined) {
          transformed[field.internalName] = this._formatDateForSharePoint(transformed[field.internalName]);
        }
      }
    }

    // Await all lookups in parallel
    const results = await Promise.all(personLookups);
    for (const { idKey, userId } of results) {
      transformed[idKey] = userId;
    }
    const multiResults = await Promise.all(multiPersonLookups);
    for (const { idKey, userIds } of multiResults) {
      transformed[idKey] = { results: userIds };
    }

    // Transform multichoice fields to SharePoint format
    return this._transformMultichoiceFieldsForSharePoint(transformed);
  }

  /**
   * Obtain list item type. Use `getListItemType` instead for cached value.
   * @protected
   * @example TBC
   * @returns {Promise<string>}
   */
  protected async getListItemTypeEx(): Promise<string> {
    // API path is now relative
    const api = `/_api/web/lists/GetByTitle('${this.list}')?$select=ListItemEntityTypeFullName`;
    const f = await this._serviceWorkerFetchWrapper(api, {
      Accept: "application/json;odata=verbose",
    });
    let data = await f.json();
    return data.d.ListItemEntityTypeFullName;
  }

  /**
   * Obtain list item type, used in creation of items.
   * Caches the result to avoid repeated API calls.
   * If multiple calls happen concurrently, they will all await the same Promise.
   * @example TBC
   * @returns {Promise<string>}
   */
  async getListItemType(): Promise<string> {
    // If there's already a Promise in flight, await it
    if (this.itemType) {
      return await this.itemType;
    }
    
    // Create a new Promise, cache it, and await it
    this.itemType = this.getListItemTypeEx();
    return await this.itemType;
  }

  /*
   * Get a list items
   * @param {number} itemId - Sharepoint Item ID of item to get
   * @example await storm.getItem(1)
   * @docs https://learn.microsoft.com/en-us/sharepoint/dev/sp-add-ins/working-with-lists-and-list-items-with-rest
   * @returns {Promise<any>}
   */
  async getItem(
    itemId: number,
    odata?: ODataQuery<TSchema>
  ): Promise<InferGetterFromSharePointSchema<TSchema>> {
    let oDataString: string = odataQueryToString(odata, this.schema);

    // API path is now relative
    const api = `/_api/web/lists/GetByTitle('${this.list}')/items(${itemId})${oDataString}`;
    const f = await this._serviceWorkerFetchWrapper(api, {
      Accept: "application/json;odata=verbose",
    });
    const data = (await f.json()).d;
    const transformed = this._transformGetterDataFromSharePoint(data);
    if (!this.itemCache.has(itemId)) this.itemCache.set(itemId, transformed);
    return transformed;
  }

  /**
   * Items are cached whenever getItem is called. `getItem` will always return the newest item from the database, but `getCachedItem` will return the first cached item.
   * This is not only useful for performance, but further utilised in determining changes made on an item since the last time it was fetched. And thus determining
   * potential conflicts with the current user's changes on that same item.
   * If the item is not cached, it will call `getItem` to fetch and cache the item.
   * @param itemId - SharePoint Item ID of item to get
   * @returns
   * @example ```
   * let initialData = await storm.getCachedItem(1)
   * let newData = await storm.getItem(1)
   * let changesOnDatabase = compare(initialData, newData)
   * let changesByUser = compare(initialData, userData)
   * let conflicts = getConflicts(changesOnDatabase, changesByUser)
   * ```
   */
  async getCachedItem(
    itemId: number
  ): Promise<InferGetterFromSharePointSchema<TSchema> | undefined> {
    if (this.itemCache.has(itemId)) {
      const cached = this.itemCache.get(itemId);
      // Ensure cached item is transformed (in case it was cached before transformation was added)
      return cached ? this._transformGetterDataFromSharePoint(cached) : undefined;
    }

    // If not cached this will fetch and cache the item
    return await this.getItem(itemId);
  }

  /**
   * Remove an item from the cache.
   * Typically an item will be removed from the cache after we have confirmation from the database that it has been updated by the user.
   * @param itemId - SharePoint Item ID of item to remove from cache
   * @return - Returns `true` if the item was successfully removed from the cache, `false` if it was not found.
   * @see getCachedItem for more information on caching.
   * @example storm.removeCachedItem(1)
   */
  removeCachedItem(itemId: number): boolean {
    if (this.itemCache.has(itemId)) {
      this.itemCache.delete(itemId);
      return true;
    }
    return false;
  }

  /*
   * Update a list item
   * @param {string | number} itemID - Sharepoint Item ID of item to update
   * @param {Record<string, any>} data  - FieldName: Value list to update
   * @remark Including values of 'undefined' will automatically be excluded from 'JSON.stringify'.
   * @example await storm.setItem(1, {"County":"Staffordshire"})
   * @returns {Promise<Response>}
   */
  async setItem(
    itemID: string | number,
    data: InferSetterFromSharePointSchema<TSchema>
  ): Promise<Response> {
    let itemType = await this.getListItemType();
    const transformedData = await this._transformSetterDataForSharePoint(data);
    transformedData["__metadata"] = { type: itemType };
    console.log({ caller: "setItem", itemID, data: transformedData });

    // API path is now relative
    const api = `/_api/web/lists/GetByTitle('${this.list}')/items(${itemID})`;
    const f = await this._serviceWorkerFetchWrapper(
      api,
      {
        Accept: "application/json;odata=verbose",
        "Content-Type": "application/json;odata=verbose",
        "If-Match": "*",
        "X-HTTP-Method": "MERGE",
      },
      "POST", // Use POST for MERGE operation with X-HTTP-Method header
      transformedData
    );
    return f;
  }

  /*
   * Add a list item
   * @param {Record<string, any>} data  - Data to add
   * @param {string} [itemType] - Optional item type, will be fetched if not provided
   * @example await storm.addItem({"County":"Staffordshire", ...})
   * @returns {Promise<number>} - Returns the ID of the newly created item
   */
  async addItem(
    data: InferCreatorFromSharePointSchema<TSchema>,
    itemType?: string
  ): Promise<number> {
    if (!itemType) itemType = await this.getListItemType();
    const transformedData = await this._transformSetterDataForSharePoint(data);
    transformedData["__metadata"] = { type: itemType };
    console.log({ caller: "addItem", itemType, data: transformedData });

    // API path is now relative
    const api = `/_api/web/lists/GetByTitle('${this.list}')/items`;
    const f = await this._serviceWorkerFetchWrapper(
      api,
      {
        Accept: "application/json;odata=verbose",
        "Content-Type": "application/json;odata=verbose",
      },
      "POST",
      transformedData
    );
    let json = await f.json();
    if (!json || !json.d || !json.d.Id) {
      throw new Error(
        `Failed to add item to list '${this.list}': ${JSON.stringify(json)}`
      );
    }

    return json.d.Id;
  }

  /*
   * Get all list items, with optional client-side filtering.
   * Handles pagination to retrieve all matching items and applies the filter locally.
   * This implementation pipelines fetches to overlap network requests with local processing.
   * @param odata - Definition of a server side filter, or any additional instructions required. It should
   *                be noted that these heavily rely on indexed columns with lists > 5000 items. In these such cases
   *                either use no odata filter with a post processing filter, or use an inplace query.
   * @param filter - An optional client-side filter function.
   * It receives each item and should return true to include the item.
   * @example await myList.getItems()
   * @example await myList.getItems((item) => item.RiskID === 1 && item.Status != 'Completed')
   * @docs https://learn.microsoft.com/en-us/sharepoint/dev/sp-add-ins/working-with-lists-and-list-items-with-rest
   * @returns {Promise<any[]>}
   */
  async getItems(
    odata?: ODataQuery<TSchema>,
    filter?: (item: InferGetterFromSharePointSchema<TSchema>) => boolean
  ): Promise<InferGetterFromSharePointSchema<TSchema>[]> {
    const oDataString = odataQueryToString(odata, this.schema);
    const initialApi = `/_api/web/lists/GetByTitle('${this.list}')/items${oDataString}`;
    const allFilteredResults: any[] = [];

    // Start the first fetch
    let currentFetchPromise: Promise<Response> =
      this._serviceWorkerFetchWrapper(initialApi, {
        Accept: "application/json;odata=verbose",
      });

    while (true) {
      // Await the current fetch promise to get the response
      const f = await currentFetchPromise;
      const data = (await f.json()).d;
      const currentResults: any[] = data.results;

      // Determine the next API path for pagination *before* processing current results
      // and start the next fetch if there's a __next link.
      let nextApi: string = "";
      let nextFetchPromise: Promise<Response> | null = null;
      if (data.__next) {
        nextApi = data.__next.substring(this.site.length);
        nextFetchPromise = this._serviceWorkerFetchWrapper(nextApi, {
          Accept: "application/json;odata=verbose",
        });
      }

      // Transform multichoice fields for each result
      const transformedResults = currentResults.map((item) =>
        this._transformGetterDataFromSharePoint(item)
      );

      // Apply client-side filter if provided and concatenate results
      if (filter) {
        allFilteredResults.push(...transformedResults.filter(filter));
      } else {
        allFilteredResults.push(...transformedResults);
      }

      // If there's no next page, break the loop
      if (!nextFetchPromise) {
        break;
      }

      // Set the next fetch promise as the current one for the next iteration
      currentFetchPromise = nextFetchPromise;
    }

    return allFilteredResults;
  }

  /*
   * Get items using an inplace Query
   * @param {string} inplaceQuery - The inplace search query string
   * @example
   * @returns {Promise<any[]>}
   */
  async getItemsWithInplaceQuery(
    inplaceQuery: string
  ): Promise<InferGetterFromSharePointSchema<TSchema>[]> {
    // API path is now relative
    const api = `/_api/web/lists/GetByTitle('${this.list}')/RenderListDataAsStream?InplaceSearchQuery=${inplaceQuery}`;
    let data = await this._serviceWorkerFetchWrapper(
      api,
      { "content-type": "application/json;odata=verbose" },
      "POST"
    );
    let json = await data.json();
    const rows = json["Row"] || [];
    // Transform multichoice fields for each result
    return rows.map((item: any) => this._transformGetterDataFromSharePoint(item));
  }

  /*
   * Update multiple items in a batch query
   * @param {IChange[]}     batch      - List of items to change
   * @param {string | null} [changeSetId=null] - UUID of the change set. By default a random UUID is generated.
   * @param {string | null} [batchUuid=null]   - UUID of the batch. By default a random UUID is generated.
   * @type IChange = {id: number, data: object} - data = object of  fields to change and values to change them to.
   * @remark https://sharepoint.stackexchange.com/questions/234766/batch-update-create-list-items-using-rest-api-in-sharepoint-2013
   * @remark https://learn.microsoft.com/en-us/sharepoint/dev/sp-add-ins/make-batch-requests-with-the-rest-apis
   * @example ```
   * await db.setItemsBatch([
   * {id: 8369, data:{NonInfraOperationalArea: "N/A"}},
   * {id: 8370, data:{NonInfraOperationalArea: "Stafford"}},
   * {id: 8371, data:{NonInfraOperationalArea: "Stratford and Warwick"}},
   * ])
   * ```
   * @returns {Promise<Response[]>}
   */
  async setItemsBatch(
    batch: ISharePointListChange<TSchema>[],
    changeSetId: string | null = null,
    batchUuid: string | null = null
  ): Promise<{type: "update", id: number, httpStatus: number, data: any, payload: ISharePointListChange<TSchema>}[]> {
    if(batch.length === 0) return [];
    let itemType = await this.getListItemType();
    if (changeSetId == null) changeSetId = generateUUID();
    if (batchUuid == null) batchUuid = generateUUID();
    const updatedItems: {type: "update", id: number, httpStatus: number, data: any, payload: ISharePointListChange<TSchema>}[] = [];

    // Get the form digest for batch operations
    const formDigest = await this._getFormDigest();

    // A max of 1000 operations are allowed in a changeset; To stay well under this value, we use batches of 750.
    for (let dataChunk of splitArrayIntoChunks(batch, 750)) {
      // Workaround for issue described @ https://learn.microsoft.com/en-us/answers/questions/1383519/error-while-trying-to-bulk-update-sharepoint-onlin
      // Adding a duplicate item to the batch as a workaround.
      dataChunk.push(dataChunk[0]); // Using dataChunk instead of data to avoid shadowing outer 'data' variable

      // Create changeset
      let batchBody: string[] = [];
      batchBody.push(`--batch_${batchUuid}`);
      batchBody.push(
        `Content-Type: multipart/mixed; boundary=changeset_${changeSetId}`
      );
      batchBody.push("");

      // Transform all data in the chunk
      const transformPromises = dataChunk.map(async (item) => {
        const transformedData = await this._transformSetterDataForSharePoint(
          item.data
        );
        transformedData["__metadata"] = { type: itemType };
        return { ...item, data: transformedData };
      });
      const transformedChunk = await Promise.all(transformPromises);

      transformedChunk.forEach((item) => {
        // Add change for each item in the batch
        batchBody.push(`--changeset_${changeSetId}`);
        batchBody.push("Content-Type:application/http");
        batchBody.push("Content-Transfer-Encoding: binary");
        batchBody.push("");
        // PATCH request line now uses the relative API path
        batchBody.push(
          `PATCH ${this.site}/_api/web/lists/GetByTitle('${this.list}')/items(${item.id}) HTTP/1.1`
        );
        batchBody.push("Content-Type: application/json;odata=verbose;");
        batchBody.push("Accept: application/json");
        batchBody.push("If-Match: *");
        batchBody.push("X-HTTP-Method: MERGE");
        batchBody.push("");
        batchBody.push(JSON.stringify(item.data));
        batchBody.push("");
      });

      // End changeset
      batchBody.push(`--changeset_${changeSetId}--`);
      // Close the main batch boundary
      batchBody.push(`--batch_${batchUuid}--`);

      

      // API path for batch requests is also relative
      const api = `/_api/$batch`;
      const f = await this._serviceWorkerFetchWrapper(
        api,
        {
          "Content-Type": `multipart/mixed; boundary="batch_${batchUuid}"`,
          "X-RequestDigest": formDigest,
        },
        "POST",
        batchBody.join("\r\n"), // The body for batch is a string
        "text" // The body type is text because the body for batch is a string
      );
      
      const responseText = await f.text();

      const matches = parseODataBatchResponse(responseText);
      
      //Note: The last match is ignored because, again, it is a repeat of the first item in the batch.
      //      see https://learn.microsoft.com/en-us/answers/questions/1383519/error-while-trying-to-bulk-update-sharepoint-onlin
      for (let i = 0; i < matches.length-1; i++) {
        const response = matches[i];
        updatedItems.push({type: "update", id: batch[i].id, httpStatus: response.statusCode, data: response.json, payload: batch[i].data as any});
      };
    }
    return updatedItems.slice(0, batch.length);
  }

  async createItemsBatch(
    batch: InferCreatorFromSharePointSchema<TSchema>[],
    batchUuid: string | null = null
  ): Promise<{type: "create", id: number, httpStatus: number, data: any, payload: InferCreatorFromSharePointSchema<TSchema>}[]> {
    if(batch.length === 0) return [];
    if (batchUuid == null) batchUuid = generateUUID();
    let itemType = await this.getListItemType();
    

    // Get the form digest for batch operations
    const formDigest = await this._getFormDigest();

    const createdItems: {type: "create", id: number, httpStatus: number, data: any, payload: InferCreatorFromSharePointSchema<TSchema>}[] = [];
    for (let dataChunk of splitArrayIntoChunks(batch, 750)) {
      const batchBody: string[] = [];
      batchBody.push(`--batch_${batchUuid}`);
      batchBody.push(
        `Content-Type: multipart/mixed; boundary=changeset_${batchUuid}`
      );
      batchBody.push("");

      // Transform each item
      const transformedChunk = await Promise.all(
        dataChunk.map(async (data) => {
          const transformed = await this._transformSetterDataForSharePoint(data);
          transformed["__metadata"] = { type: itemType };
          return transformed;
        })
      );
      
      transformedChunk.forEach((data) => {
        batchBody.push(`--changeset_${batchUuid}`);
        batchBody.push("Content-Type: application/http");
        batchBody.push("Content-Transfer-Encoding: binary");
        batchBody.push("");
        batchBody.push(
          `POST ${this.site}/_api/web/lists/GetByTitle('${this.list}')/items HTTP/1.1`
        );
        batchBody.push("Content-Type: application/json;odata=verbose;");
        batchBody.push("Accept: application/json;odata=verbose");
        batchBody.push("");
        batchBody.push(JSON.stringify(data));
        batchBody.push("");
      });

      batchBody.push(`--changeset_${batchUuid}--`);
      batchBody.push(`--batch_${batchUuid}--`);

      const api = `/_api/$batch`;
      const resp = await this._serviceWorkerFetchWrapper(
        api,
        {
          "Content-Type": `multipart/mixed; boundary="batch_${batchUuid}"`,
          "X-RequestDigest": formDigest,
        },
        "POST",
        batchBody.join("\r\n"),
        "text"
      );
      const responseText = await resp.text();
      
      const matches = parseODataBatchResponse(responseText);
      for (let i = 0; i < matches.length; i++) {
        const response = matches[i];
        createdItems.push({type: "create",id: response.json?.d?.Id ?? 0, httpStatus: response.statusCode, data: response.json, payload: batch[i]});
      };
    }

    return createdItems.slice(0, batch.length);
  }

  /**
   * Batch delete items by their SharePoint IDs.
   * @param ids - Array of SharePoint item IDs to delete
   * @param changeSetId - Optional UUID for the changeset; if not provided, a new UUID will be generated
   * @param batchUuid - Optional UUID for the batch; if not provided, a new UUID will be generated
   * @returns Promise resolving to an array of Responses from each batch request
   */
  async deleteItemsBatch(
    ids: number[],
    changeSetId: string | null = null,
    batchUuid: string | null = null
  ): Promise<{type: "delete", id: number, httpStatus: number, data: any}[]> {
    if(ids.length === 0) return [];
    if (changeSetId == null) changeSetId = generateUUID();
    if (batchUuid == null) batchUuid = generateUUID();
    const deletedItems: {type: "delete", id: number, httpStatus: number, data: any}[] = [];

    // Get the form digest for batch operations
    const formDigest = await this._getFormDigest();

    // Similar rule: max 1000 ops per batch, we stay under (750 here)
    for (const idChunk of splitArrayIntoChunks(ids, 750)) {
      const batchBody: string[] = [];

      // Batch + changeset boundaries
      batchBody.push(`--batch_${batchUuid}`);
      batchBody.push(
        `Content-Type: multipart/mixed; boundary=changeset_${changeSetId}`
      );
      batchBody.push("");

      for (const id of idChunk) {
        batchBody.push(`--changeset_${changeSetId}`);
        batchBody.push("Content-Type: application/http");
        batchBody.push("Content-Transfer-Encoding: binary");
        batchBody.push("");

        // DELETE line
        batchBody.push(
          `DELETE ${this.site}/_api/web/lists/GetByTitle('${this.list}')/items(${id}) HTTP/1.1`
        );
        batchBody.push("Accept: application/json;odata=verbose");
        batchBody.push("If-Match: *"); // ensure delete even if item updated
        batchBody.push("");
      }

      // End changeset + batch
      batchBody.push(`--changeset_${changeSetId}--`);
      batchBody.push(`--batch_${batchUuid}--`);

      // Send once per chunk
      const api = `/_api/$batch`;
      const f = await this._serviceWorkerFetchWrapper(
        api,
        {
          "Content-Type": `multipart/mixed; boundary="batch_${batchUuid}"`,
          "X-RequestDigest": formDigest,
        },
        "POST",
        batchBody.join("\r\n"),
        "text"
      );
      
      const responseText = await f.text();
      const matches = parseODataBatchResponse(responseText);
      for (let i = 0; i < matches.length; i++) {
        const response = matches[i];
        deletedItems.push({type: "delete", id: ids[i], httpStatus: response.statusCode, data: response.json});
      };
    }

    // Return only the items that were actually deleted
    return deletedItems.slice(0, ids.length);
  } 
  
  /**
   * Simplifies an update payload by removing fields that haven't changed from the cached version.
   * For "create" and "delete" payloads, returns them unchanged.
   * @param payload The commit payload to simplify
   * @returns A simplified payload with only changed fields (or null if no changes)
   */
  async simplifyPayload(
    payload: ICommitPayload<TSchema>
  ): Promise<ICommitPayload<TSchema> | null> {
    if (!payload || payload.type !== "update") {
      // Create and delete payloads don't need simplification
      return payload;
    }
  
    const { id, data } = payload.payload;
    
    // Get the cached original data
    const cached = await this.getCachedItem(id);
    
    if (!cached) {
      // If no cache exists, return the full payload
      // (This shouldn't happen for modified items, but handle gracefully)
      return payload;
    }
  
    // Compare and build only changed fields
    const changedFields: Partial<InferSetterFromSharePointSchema<TSchema>> = {};
  
    for (const key in data) {
      const currentValue = data[key];
      const cachedValue = (cached as any)[key];
      
      // Find the field schema to determine field type
      const fieldSchema = this.schema.find(f => f.internalName === key);
      
      // Normalize values for comparison based on field type
      let normalizedCurrentValue = currentValue;
      let normalizedCachedValue = cachedValue;
      
      // For date fields, normalize to timestamps for comparison
      if (fieldSchema?.type === "date") {
        normalizedCurrentValue = this._normalizeDateForComparison(currentValue);
        normalizedCachedValue = this._normalizeDateForComparison(cachedValue);
      }
  
      // Deep comparison for arrays
      if (Array.isArray(normalizedCurrentValue) && Array.isArray(normalizedCachedValue)) {
        // Sort arrays for comparison (order might differ but content is same)
        // Convert to strings for comparison to handle nested arrays/objects
        const currentSorted = JSON.stringify([...normalizedCurrentValue].sort());
        const cachedSorted = JSON.stringify([...normalizedCachedValue].sort());
        if (currentSorted !== cachedSorted) {
          changedFields[key] = currentValue;
        }
      }
      // Handle case where one is array and other is not
      else if (Array.isArray(normalizedCurrentValue) !== Array.isArray(normalizedCachedValue)) {
        changedFields[key] = currentValue;
      }
      // Handle null/undefined comparisons
      else if (normalizedCurrentValue === null || normalizedCurrentValue === undefined) {
        if (normalizedCachedValue !== normalizedCurrentValue) {
          changedFields[key] = currentValue;
        }
      }
      // Regular comparison for primitives (using normalized values for date fields)
      else if (normalizedCurrentValue !== normalizedCachedValue) {
        changedFields[key] = currentValue;
      }
    }
  
    // If no fields changed, return null (no update needed)
    if (Object.keys(changedFields).length === 0) {
      return null;
    }
  
    // Return simplified payload with only changed fields
    return {
      type: "update",
      payload: {
        id,
        data: changedFields,
      },
    } as ICommitPayload<TSchema>;
  }

  /**
   * Normalizes date values to timestamps for comparison.
   * Handles SharePoint date format (YYYY-MM-DD hh:mm:ss) and ISO format (YYYY-MM-DDThh:mm:ss.sssZ).
   * @param value - The value to normalize (can be a date string, Date object, or any other value)
   * @returns The timestamp (number) if it's a date, otherwise the original value
   */
  private _normalizeDateForComparison(value: any): any {
    if (value === null || value === undefined) {
      return value;
    }
    
    // If it's already a Date object, return its timestamp
    if (value instanceof Date) {
      return value.getTime();
    }
    
    // If it's a number (already a timestamp), return as is
    if (typeof value === "number") {
      return value;
    }
    
    // If it's a string, try to parse it as a date
    if (typeof value === "string" && value.trim() !== "") {
      // Try parsing as ISO format first (JavaScript standard)
      // This handles formats like "2025-07-07T23:00:00.000Z" or "2025-07-07T23:00:00"
      const isoDate = new Date(value);
      if (!isNaN(isoDate.getTime())) {
        return isoDate.getTime();
      }
      
      // Try parsing SharePoint format: "YYYY-MM-DD hh:mm:ss" or "YYYY-MM-DD hh:mm:ss.sss"
      // SharePoint dates might be in UTC or local time, so try both
      const sharePointDateMatch = value.match(/^(\d{4})-(\d{2})-(\d{2})\s+(\d{2}):(\d{2}):(\d{2})(?:\.(\d+))?/);
      if (sharePointDateMatch) {
        const [, year, month, day, hour, minute, second, milliseconds] = sharePointDateMatch;
        const ms = milliseconds ? parseInt(milliseconds.padEnd(3, '0').substring(0, 3)) : 0;
        
        // Try parsing as UTC first (most common for SharePoint)
        const utcDate = Date.UTC(
          parseInt(year),
          parseInt(month) - 1, // Month is 0-indexed
          parseInt(day),
          parseInt(hour),
          parseInt(minute),
          parseInt(second),
          ms
        );
        if (!isNaN(utcDate)) {
          return utcDate;
        }
        
        // Fallback to local time parsing
        const localDate = new Date(
          parseInt(year),
          parseInt(month) - 1,
          parseInt(day),
          parseInt(hour),
          parseInt(minute),
          parseInt(second),
          ms
        );
        if (!isNaN(localDate.getTime())) {
          return localDate.getTime();
        }
      }
    }
    
    // If we can't parse it as a date, return the original value
    return value;
  }

  /**
   * Checks for conflicts between cached data, user changes, and current database state.
   * A conflict exists when:
   * - The user changed a field (oldValue !== newValue)
   * - The database also changed that field (dbValue !== oldValue)
   * - The user's change differs from the database change (newValue !== dbValue)
   * 
   * @param payloads - Array of commit payloads to check for conflicts
   * @returns Array of conflicts found, empty if no conflicts
   * @example
   * ```typescript
   * const payload = await risk.toPayload();
   * const conflicts = await TorrentDB.Risks.checkConflicts([payload]);
   * if (conflicts.length > 0) {
   *   // Handle conflicts
   *   conflicts.forEach(conflict => {
   *     conflict.resolve(conflict.conflictValue); // Accept database value
   *   });
   * }
   * ```
   */
  async checkConflicts(
    payloads: ICommitPayload<TSchema>[]
  ): Promise<IConflict<any>[]> {
    const conflicts: IConflict<any>[] = [];
    const validPayloads = payloads.filter((p): p is Exclude<ICommitPayload<TSchema>, null> => p !== null);
    
    // Only check conflicts for "update" payloads (creates/deletes don't have conflicts)
    const updatePayloads = validPayloads.filter((p): p is Extract<ICommitPayload<TSchema>, { type: "update" }> => 
      p.type === "update"
    );
    
    for (const payload of updatePayloads) {
      const { id, data } = payload.payload;
      
      // Get cached (old) and current (conflict) values
      const cachedItem = await this.getCachedItem(id);
      const currentDBItem = await this.getItem(id);
      
      if (!cachedItem) {
        // No cache means this is a new item or cache was cleared - skip conflict checking
        continue;
      }
      
      // Check each field in the payload data for conflicts
      for (const fieldName in data) {
        const newValue = data[fieldName];
        
        // Handle Person field name mapping: "ClosedByEmail" -> "ClosedBy"
        // The payload uses Email suffix, but schema and DB use base field name
        let baseFieldName = fieldName;
        let isPersonEmailField = false;
        if (fieldName.endsWith("Email")) {
          baseFieldName = fieldName.slice(0, -5); // Remove "Email" suffix
          isPersonEmailField = true;
        }
        
        const oldValue = (cachedItem as any)[baseFieldName];
        const dbValue = (currentDBItem as any)[baseFieldName];
        
        // Find the field schema to determine field type (use base field name)
        const fieldSchema = this.schema.find(f => f.internalName === baseFieldName);
        
        // Normalize values for comparison based on field type
        let normalizedOldValue = oldValue;
        let normalizedNewValue = newValue;
        let normalizedDbValue = dbValue;
        
        // For date fields, normalize to timestamps for comparison
        if (fieldSchema?.type === "date") {
          normalizedOldValue = this._normalizeDateForComparison(oldValue);
          normalizedNewValue = this._normalizeDateForComparison(newValue);
          normalizedDbValue = this._normalizeDateForComparison(dbValue);
        } else if (fieldSchema?.type === "person") {
          // Person fields: payload has Email string; cache/DB have Person objects.
          // Normalize all to Email | null for like-for-like comparison, and treat
          // undefined vs null as equivalent (avoids false conflicts when cache
          // omits field (undefined) but DB has null).
          normalizedOldValue = (oldValue as SharePointPerson | null | undefined)?.Email ?? null;
          normalizedNewValue = newValue ?? null; // Already Email from payload
          normalizedDbValue = (dbValue as SharePointPerson | null | undefined)?.Email ?? null;
        }
        
        // Treat undefined and null as equivalent "empty" for conflict detection.
        // Cache may omit fields (undefined) while DB stores null, causing
        // dbValue !== oldValue (null !== undefined) and false-positive conflicts.
        const empty = (v: unknown) => (v == null ? null : v);
        const o = empty(normalizedOldValue);
        const n = empty(normalizedNewValue);
        const d = empty(normalizedDbValue);
        
        // Check if conflict exists: user changed (o !== n), DB changed (d !== o), and they differ (n !== d)
        const hasConflict = o !== n && n !== d && d !== o;
        
        if (hasConflict) {
          // Values for the conflict object (use normalized Email for person fields)
          let conflictOldValue: any = normalizedOldValue;
          let conflictNewValue: any = normalizedNewValue;
          let conflictDbValue: any = normalizedDbValue;
          
          if (fieldSchema?.type === "person") {
            conflictOldValue = (oldValue as SharePointPerson | null)?.Email || null;
            conflictNewValue = newValue; // Already in Email format from payload
            conflictDbValue = (dbValue as SharePointPerson | null)?.Email || null;
          } else if (fieldSchema?.type === "multiperson") {
            conflictOldValue = Array.isArray(oldValue) 
              ? (oldValue as SharePointPerson[]).map(p => p.Email).filter(Boolean)
              : null;
            conflictNewValue = newValue; // Already in Email array format from payload
            conflictDbValue = Array.isArray(dbValue)
              ? (dbValue as SharePointPerson[]).map(p => p.Email).filter(Boolean)
              : null;
          }
          
          // Create resolve function that removes the cache
          // Note: The domain class (e.g., Risk) will wrap this to also update its instance properties
          // We only remove the cache here - updating the domain object is the domain class's responsibility
          const resolveConflict = (value: any) => {
            // Remove cached item to ensure fresh data is fetched next time
            // This prevents the conflict from being detected again
            this.removeCachedItem(id);
          };
          
          conflicts.push({
            itemId: id,
            fieldName,
            oldValue: conflictOldValue,
            newValue: conflictNewValue,
            conflictValue: conflictDbValue,
            resolve: resolveConflict,
          });
        }
      }
    }
    
    return conflicts;
  }

  async processCommitPayloads(payloads: ICommitPayload<TSchema>[]): Promise<any[]> {
    let validPayloads = payloads.filter(p => p !== null);
    let updates = validPayloads.filter(p => p.type === "update").map(p => p.payload);
    let creates = validPayloads.filter(p => p.type === "create").map(p => p.payload);
    let deletes = validPayloads.filter(p => p.type === "delete").map(p => p.payload.id);
    return await Promise.all([
      this.setItemsBatch(updates),
      this.createItemsBatch(creates),
      this.deleteItemsBatch(deletes)
    ] as const)
  }
}

export type ICommitPayload<TSchema extends readonly SharePointFieldSchema[]> = 
  | { type: "create"; payload: InferCreatorFromSharePointSchema<TSchema>; } 
  | { type: "update"; payload: ISharePointListChange<TSchema>; } 
  | { type: "delete"; payload: { id: number; }; }
  | null;

/**
 * Commit payload for document library operations.
 * Extends ICommitPayload with file-specific operations.
 */
export type IDocumentLibraryCommitPayload<TSchema extends readonly SharePointFieldSchema[]> =
  | { 
      type: "create"; 
      payload: InferCreatorFromSharePointSchema<TSchema>;
      folderPath?: string;
      file: File | Blob;
      fileName?: string;
    }
  | { 
      type: "update"; 
      payload: ISharePointListChange<TSchema>;
      file?: File | Blob;
      folderPath?: string;
      fileName?: string;
    }
  | { type: "delete"; payload: { id: number; }; }
  | null;

/**
 * SharePoint Document Library class.
 * Extends SharePointList to provide file management capabilities while reusing
 * all list field operations from the base class.
 * 
 * Document libraries use the same SharePoint list APIs for field updates,
 * but have separate APIs for file upload, download, and file operations.
 * 
 * @example
 * const docLib = new SharePointDocumentLibrary(
 *   "https://site.sharepoint.com/sites/SiteName",
 *   "Documents",
 *   schema
 * );
 * 
 * // Upload a new file and set metadata
 * await docLib.processCommitPayloads([{
 *   type: "create",
 *   payload: { Title: "My Document" },
 *   folderPath: "SubFolder",
 *   fileName: "document.pdf",
 *   file: fileBlob
 * }]);
 */
export class SharePointDocumentLibrary<TSchema extends readonly SharePointFieldSchema[]> 
  extends SharePointList<AppendFileToSchema<TSchema>> {
  
  /**
   * @constructor
   * @param {string} site - The base URL of the SharePoint site
   * @param {string} listName - The name of the SharePoint document library
   * @param {TSchema} baseSchema - The base schema (File field will be automatically added)
   * @example new SharePointDocumentLibrary("https://site.sharepoint.com/sites/SiteName", "Documents", schema)
   */
  constructor(site: string, listName: string, baseSchema: TSchema) {
    // Check if File field already exists in the schema
    const hasFileField = baseSchema.some(
      (field) => field.internalName === "File" && field.type === "file"
    );
    
    // Create File field schema
    const fileField: FileFieldSchema<typeof DefaultFileSelectedProps> = {
      internalName: "File",
      type: "file",
      selectProps: DefaultFileSelectedProps,
    };
    
    // Append File field to schema if it doesn't already exist
    // Type assertion is safe because AppendFileToSchema ensures File is present
    const extendedSchema = hasFileField 
      ? baseSchema as unknown as AppendFileToSchema<TSchema>
      : ([...baseSchema, fileField] as const) as AppendFileToSchema<TSchema>;
    
    super(site, listName, extendedSchema);
  }
  
  /**
   * Get a list item (overridden to ensure proper type inference with File field)
   * @param itemId - SharePoint Item ID of item to get
   * @param odata - Optional OData query
   * @returns The item with all schema fields including File
   */
  override async getItem(
    itemId: number,
    odata?: ODataQuery<AppendFileToSchema<TSchema>>
  ): Promise<DocumentLibraryResult<TSchema>> {
    const result = await super.getItem(itemId, odata);
    return result as unknown as DocumentLibraryResult<TSchema>;
  }
  
  /**
   * Get all list items (overridden to ensure proper type inference with File field)
   * @param odata - Optional OData query
   * @param filter - Optional client-side filter function
   * @returns Array of items with all schema fields including File
   */
  override async getItems(
    odata?: ODataQuery<AppendFileToSchema<TSchema>>,
    filter?: (item: DocumentLibraryResult<TSchema>) => boolean
  ): Promise<DocumentLibraryResult<TSchema>[]> {
    const result = await super.getItems(odata, filter as any);
    return result as unknown as DocumentLibraryResult<TSchema>[];
  }
  
  /**
   * Get cached item (overridden to ensure proper type inference with File field)
   * @param itemId - SharePoint Item ID of item to get
   * @returns The cached item with all schema fields including File
   */
  override async getCachedItem(
    itemId: number
  ): Promise<DocumentLibraryResult<TSchema> | undefined> {
    const result = await super.getCachedItem(itemId);
    return result as unknown as DocumentLibraryResult<TSchema> | undefined;
  }
  
  /**
   * Upload a file to a document library folder.
   * @param file - The file content as File or Blob 
   * @param fileName - The name of the file to upload
   * @param folderPath - The folder path relative to the document library root (e.g., "SubFolder" or "SubFolder/Nested", or "/" for root).
   *                     Do not include the list name - it should be relative to the document library root only.
   * @param overwrite - Whether to overwrite if file exists (default: true)
   * @returns Promise resolving to the list item ID of the uploaded file
   */
  async fileCreateItem(
    file: File | Blob,
    fileName?: string,
    folderPath?: string,
    overwrite: boolean = true
  ): Promise<number> {   
    //If the file is a File object, and fileName is not provided, use the file's name
    if (file instanceof File && !fileName) {
      fileName = file.name;
    }
    //If the file is a Blob object (but not a File), and fileName is not provided, throw an error
    if (file instanceof Blob && !(file instanceof File) && !fileName) {
      throw new Error("fileName is required");
    }

    // Normalize the folder path - use / for root, /SubFolder for subfolders
    // The service worker expects paths relative to the document library root, not including list name
    let folderPathForServiceWorker: string;
    if (!folderPath) {
      // If no folder path provided, use root
      folderPathForServiceWorker = '/';
    } else {
      // Remove leading/trailing slashes and ensure it starts with /
      const cleanFolderPath = folderPath.replace(/^\/+|\/+$/g, '');
      folderPathForServiceWorker = cleanFolderPath ? `/${cleanFolderPath}` : '/';
    }
    
    // Use the new file-create operation
    let res: Response;
    try {
      res = await this._serviceWorkerFetchWrapperFile(
        file,
        folderPathForServiceWorker,
        fileName!,
        "file-create"
      );
    } catch (error) {
      console.error("fileCreateItem: error in _serviceWorkerFetchWrapperFile", error);
      throw error;
    }

    // The service worker should return the SharePoint response in the same format
    const result = await res.json();
    // SharePoint returns the file info in d object, which includes ListItemAllFields.Id
    const itemId = result.ItemId;
    
    if (!itemId) {
      throw new Error(
        `Failed to get item ID after file upload: ${JSON.stringify(result)}`
      );
    }

    return itemId;
  }

  /**
   * Get the content of a file from the document library.
   * @param itemId - The list item ID of the file
   * @returns Promise resolving to a File object with the file content
   */
  async fileGetContent(itemId: number): Promise<File> {
    const item = await this.getItem(itemId);
    
    // Get the server-relative URL and file name
    const serverRelativeUrl = item.File.ServerRelativeUrl;
    const fileName = item.File.Name;
    
    // Escape single quotes in the URL (SharePoint requires doubling them)
    const escapedUrl = serverRelativeUrl.replace(/'/g, "''");
    
    // Use GetFileByServerRelativeUrl with $value endpoint to get binary content
    const api = `/_api/web/getfilebyserverrelativeurl('${escapedUrl}')/$value`;
    
    // Make the request - $value endpoint returns binary content, not JSON
    const response = await this._serviceWorkerFetchWrapper(
      api,
      {
        Accept: "*/*", // Accept any content type for binary files
      },
      "GET"
    );
    
    // Get the blob from the response
    const blob = await response.blob();
    
    // Convert blob to File object with the original filename
    const file = new File([blob], fileName, { type: blob.type });
    
    return file;
  }

  /**
   * Update the content of an existing file.
   * @param itemId - The list item ID of the file
   * @param file - The new file content as File or Blob
   * @returns Promise resolving to the Response
   */
  async fileUpdateContent(
    itemId: number,
    file: File | Blob
  ): Promise<Response> {
    
    // First get the file to find its server-relative URL
    // FileRef, FileDirRef, FileLeafRef are SharePoint system fields not in schema
    const item = await this.getItem(itemId)
    const fileName = item.File.Name
    // Get the server relative URL of the file e.g. "/sites/Asset/Planning/Risks/Ops/Data/Attachments/19d0ade1-7049-4884-99bb-ba7a1153b87f.txt"
    const serverRelativeURL = item.File.ServerRelativeUrl
    // Obtain the site name in the format `/sites/...` (from https://site.sharepoint.com/sites/Subsite/SiteName)
    const siteURL = /sites\/.+/.exec(this.site)?.[0];

    
    // Use a regex to safely extract the relative folder path between list and file name
    let folderPath = "";
    if (typeof serverRelativeURL === "string" && typeof fileName === "string") {
      // Build the regex: /sites/.../<listname>/(folderpath/)?fileName
      // We want to capture the (optional) folder path part, if any
      // Example: /sites/Asset/Planning/Risks/Ops/Data/Attachments/SubFolder/thing.txt
      //            ^siteURL                 ^list        ^folderPath   ^fileName
      // Regex:   ^/(sites/.*?)/list/(.*?/)??fileName$
      // Replace any special chars in the list name and file name to avoid regex injection
      const escapedList = this.list.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
      const escapedFileName = fileName.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
      // Only match if siteURL is valid
      if (siteURL) {
        const regex = new RegExp(
          `^\\/${siteURL}\\/${escapedList}(\\/.*)${escapedFileName}$`
        );
        const match = serverRelativeURL.match(regex);
        if (match) {
          folderPath = match[1]
        }
      }
    }
    console.log({fileName, serverRelativeURL, folderPath, siteURL});

    // Use the new file-update operation
    return await this._serviceWorkerFetchWrapperFile(
      file,
      folderPath || "/",
      fileName,
      "file-update"
    );
  }

  /**
   * Move or rename a file in the document library.
   * @param itemId - The list item ID of the file to move
   * @param newFolderPath - Optional new folder path relative to the document library (e.g., "SubFolder" or "SubFolder/Nested"). 
   *                        The list name is automatically prepended, so you don't need to include it.
   *                        If not provided, the file will remain in its current folder.
   * @param newFileName - Optional new file name (if not provided, keeps current name)
   * @returns Promise resolving to the Response
   */
  async fileMove(
    itemId: number,
    newFolderPath?: string,
    newFileName?: string
  ): Promise<Response> {
    // Get current file info
    const item = await this.getItem(itemId);
    
    // Get current file info
    const currentServerRelativeUrl = item.File.ServerRelativeUrl;
    const currentFileName = item.File.Name;
    
    // Use provided filename or keep current filename
    const targetFileName = newFileName || currentFileName;
    
    // Build the new server-relative URL
    // Current URL format: /sites/.../ListName/currentFolder/file.txt or /sites/.../ListName/file.txt (root)
    // Extract the base path (site + list name + trailing slash)
    const listNamePattern = `/${this.list}`;
    const listNameIndex = currentServerRelativeUrl.indexOf(listNamePattern);
    
    if (listNameIndex === -1) {
      throw new Error(`Cannot find list name '${this.list}' in ServerRelativeUrl: ${currentServerRelativeUrl}`);
    }
    
    // Extract base path: everything up to and including the list name, plus a trailing slash
    // This gives us: /sites/.../ListName/
    const basePath = currentServerRelativeUrl.substring(0, listNameIndex + listNamePattern.length) + '/';
    
    // If newFolderPath is not provided, extract the current folder path from the file's URL
    let folderPathToUse: string;
    if (newFolderPath === undefined) {
      // Extract current folder path: everything between the list name and the filename
      const pathAfterList = currentServerRelativeUrl.substring(listNameIndex + listNamePattern.length + 1);
      const fileNameIndex = pathAfterList.lastIndexOf('/');
      if (fileNameIndex === -1) {
        // File is in root folder
        folderPathToUse = '';
      } else {
        // Extract folder path (remove leading/trailing slashes)
        folderPathToUse = pathAfterList.substring(0, fileNameIndex).replace(/^\/+|\/+$/g, '');
      }
    } else {
      // Normalize newFolderPath (remove leading/trailing slashes)
      folderPathToUse = newFolderPath.replace(/^\/+|\/+$/g, '');
    }
    
    const normalizedFolderPath = folderPathToUse;
    
    // Build the destination folder path (without filename) to check if it exists
    const destinationFolderPath = normalizedFolderPath 
      ? `${basePath}${normalizedFolderPath}`
      : basePath.slice(0, -1); // Remove trailing slash for root folder check
    
    // Check if the destination folder exists before attempting to move
    // Only check if we're moving to a subfolder (not root)
    if (normalizedFolderPath) {
      try {
        const escapedFolderPath = destinationFolderPath.replace(/'/g, "''");
        const folderCheckApi = `/_api/web/getfolderbyserverrelativeurl('${escapedFolderPath}')`;
        await this._serviceWorkerFetchWrapper(
          folderCheckApi,
          {
            Accept: "application/json;odata=verbose",
          },
          "GET"
        );
      } catch (error: any) {
        // If folder doesn't exist, throw a descriptive error
        if (error?.data?.status === 404 || error?.message?.includes('404')) {
          throw new Error(
            `Destination folder does not exist: ${destinationFolderPath}. ` +
            `Please create the folder '${normalizedFolderPath}' in the document library '${this.list}' before moving the file.`
          );
        }
        // Re-throw other errors
        throw error;
      }
    }
    
    // Build the new server-relative URL
    const newFileRef = normalizedFolderPath 
      ? `${basePath}${normalizedFolderPath}/${targetFileName}`
      : `${basePath}${targetFileName}`;

    // For document library files, use GetFileByServerRelativeUrl instead of list item endpoint
    // This is the recommended approach for file operations in document libraries
    // Note: Single quotes in the ServerRelativeUrl need to be escaped by doubling them
    const escapedCurrentUrl = currentServerRelativeUrl.replace(/'/g, "''");
    const escapedNewUrl = newFileRef.replace(/'/g, "''");
    const api = `/_api/web/getfilebyserverrelativeurl('${escapedCurrentUrl}')/moveto(newurl='${escapedNewUrl}',flags=1)`;
    
    // SharePoint MoveTo requires POST with Content-Type header and empty JSON body {}
    // The "end of input stream" error occurs when Content-Type is set but body is missing
    return await this._serviceWorkerFetchWrapper(
      api,
      {
        Accept: "application/json;odata=verbose",
        "Content-Type": "application/json;odata=verbose",
        "IF-MATCH": "*", // Required for file operations
      },
      "POST",
      {} // Empty JSON object body - SharePoint MoveTo requires this
    );
  }

  /**
   * Rename a file (convenience method that moves within the same folder).
   * @param itemId - The list item ID of the file
   * @param newFileName - The new file name
   * @returns Promise resolving to the Response
   */
  async fileRename(itemId: number, newFileName: string): Promise<Response> {
    // Use fileMove with undefined folder path to keep file in current folder
    return this.fileMove(itemId, undefined, newFileName);
  }

  /**
   * Process commit payloads for document library operations.
   * Handles file uploads/updates before processing standard list field operations.
   * @param payloads - Array of document library commit payloads
   */
  async processCommitPayloads(
    payloads: IDocumentLibraryCommitPayload<TSchema>[]
  ): Promise<any[]> {
    const validPayloads = payloads.filter((p) => p !== null);
    
    // Separate payloads by type and file operations
    //New files to be created
    const createsWithFiles: Array<{
      type: "create";
      payload: InferCreatorFromSharePointSchema<TSchema>;
      folderPath?: string;
      file: File | Blob;
      fileName?: string;
    }> = validPayloads
      .filter((p): p is Extract<IDocumentLibraryCommitPayload<TSchema>, { type: "create" }> => 
        p?.type === "create" && !!p.file && !!p.fileName
      )
    
    //Files to be updated
    const updatesWithFiles: Array<{
      itemId: number;
      payload: InferSetterFromSharePointSchema<TSchema>;
      file?: File | Blob;
      folderPath?: string;
      fileName?: string;
    }> = [];
    
    //Standard field updates
    const updatesWithoutFiles: ISharePointListChange<TSchema>[] = [];

    //Files to be deleted
    const deletes: number[] = [];

    // Process files to be updated
    for (const p of validPayloads.filter((p) => p?.type === "update")) {
      if (!p || p.type !== "update") continue;
      
      if (p.file || p.folderPath || p.fileName) {
        // Handle file operations
        updatesWithFiles.push({
          itemId: p.payload.id,
          payload: p.payload.data,
          file: p.file,
          folderPath: p.folderPath,
          fileName: p.fileName,
        });
      } else {
        // Standard field update only
        updatesWithoutFiles.push(p.payload);
      }
    }

    // Process deletes
    for (const p of validPayloads.filter((p) => p?.type === "delete")) {
      if (p && p.type === "delete") {
        deletes.push(p.payload.id);
      }
    }

    // Execute operations in parallel where possible
    const operations: {type: string, data: any, success: boolean}[] = [];

    // Upload new files and then update their metadata
    if (createsWithFiles.length > 0) {
      await (async () => {
        const uploadResults: {itemId: number, data: any}[] = await Promise.all(
          createsWithFiles.map(async (c) => {
            let itemId: number;
            try {
              // folderPath should be relative to document library root (not including list name)
              // fileCreateItem expects paths like "/SubFolder" or "/" for root
              itemId = await this.fileCreateItem(
                c.file,
                c.fileName!,
                c.folderPath,
                true
              );
              operations.push({type: "fileCreateItem", data: {itemId}, success: true})
            } catch (error) {
              operations.push({type: "fileCreateItem", data: {error}, success: false})
            }
            return {itemId, data: c.payload};
          })
        );

        // Batch update metadata for all newly uploaded files
        if (uploadResults.length > 0) {
          const metadataUpdates = uploadResults.map((r) => ({
            id: r.itemId,
            data: r.data,
          }));

          const responses = await super.setItemsBatch(metadataUpdates);
          responses.forEach(r=>{
            operations.push({type: "setItemsBatch", data: r, success: r.httpStatus >= 200 && r.httpStatus < 300})
          })
        }
      })()
      
    }

    // Update files and their metadata
    if (updatesWithFiles.length > 0) {
      await (async () => {
        // Process all file operations individually
        const fileOperationPromises: Promise<any>[] = [];
        
        updatesWithFiles.forEach((u) => {
          // Update file content if provided
          if (u.file) {
            fileOperationPromises.push(
              this.fileUpdateContent(u.itemId, u.file).then(
                (response) => {
                  operations.push({
                    type: "fileUpdateContent",
                    data: response,
                    success: response.ok
                  });
                  return response;
                }
              ).catch((error) => {
                operations.push({
                  type: "fileUpdateContent",
                  data: { error },
                  success: false
                });
                throw error;
              })
            );
          }

          // Move file if folder path or file name changed
          if (u.folderPath || u.fileName) {
            if (u.folderPath && u.fileName) {
              fileOperationPromises.push(
                this.fileMove(u.itemId, u.folderPath, u.fileName).then(
                  (response) => {
                    operations.push({
                      type: "fileMove",
                      data: response,
                      success: response.ok
                    });
                    return response;
                  }
                ).catch((error) => {
                  operations.push({
                    type: "fileMove",
                    data: { error },
                    success: false
                  });
                  throw error;
                })
              );
            } else if (u.fileName) {
              fileOperationPromises.push(
                this.fileRename(u.itemId, u.fileName).then(
                  (response) => {
                    operations.push({
                      type: "fileRename",
                      data: response,
                      success: response.ok
                    });
                    return response;
                  }
                ).catch((error) => {
                  operations.push({
                    type: "fileRename",
                    data: { error },
                    success: false
                  });
                  throw error;
                })
              );
            } else if (u.folderPath) {
              // Get current file name to keep it
              // FileLeafRef is a SharePoint system field not in schema
              fileOperationPromises.push(
                this.getItem(u.itemId, {
                  $select: ["FileLeafRef"] as any,
                }).then((item) => {
                  const currentName = (item as any).FileLeafRef || "";
                  return this.fileMove(u.itemId, u.folderPath, currentName).then(
                    (response) => {
                      operations.push({
                        type: "fileMove",
                        data: response,
                        success: response.ok
                      });
                      return response;
                    }
                  ).catch((error) => {
                    operations.push({
                      type: "fileMove",
                      data: { error },
                      success: false
                    });
                    throw error;
                  });
                })
              );
            }
          }
        });

        // Wait for all file operations to complete
        await Promise.all(fileOperationPromises);

        // Collect all field updates and batch them
        const fieldUpdates: ISharePointListChange<TSchema>[] = updatesWithFiles
          .filter((u) => Object.keys(u.payload).length > 0)
          .map((u) => ({
            id: u.itemId,
            data: u.payload,
          }));

        // Batch update all metadata fields
        if (fieldUpdates.length > 0) {
          const responses = await super.setItemsBatch(fieldUpdates);
          responses.forEach((r) => {
            operations.push({
              type: "setItemsBatch",
              data: r,
              success: r.httpStatus >= 200 && r.httpStatus < 300
            });
          });
        }
      })();
    }

    // Update items without file operations (using parent class method)
    if (updatesWithoutFiles.length > 0) {
      const responses = await super.setItemsBatch(updatesWithoutFiles);
      responses.forEach((r) => {
        operations.push({
          type: "setItemsBatch",
          data: r,
          success: r.httpStatus >= 200 && r.httpStatus < 300
        });
      });
    }

    // Delete items (using parent class method)
    if (deletes.length > 0) {
      const responses = await super.deleteItemsBatch(deletes);
      responses.forEach((r) => {
        operations.push({
          type: "deleteItemsBatch",
          data: r,
          success: r.httpStatus >= 200 && r.httpStatus < 300
        });
      });
    }

    return operations;
  }
}



