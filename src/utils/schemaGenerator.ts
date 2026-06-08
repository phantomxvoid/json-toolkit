type SchemaNode = Record<string, unknown>;

function inferSchema(value: unknown): SchemaNode {
  if (value === null) return { type: "null" };

  if (Array.isArray(value)) {
    if (value.length === 0) return { type: "array", items: {} };
    const itemSchemas = value.map(inferSchema);
    return { type: "array", items: mergeSchemas(itemSchemas) };
  }

  if (typeof value === "object") {
    const obj = value as Record<string, unknown>;
    const properties: SchemaNode = {};
    const required: string[] = [];
    for (const [k, v] of Object.entries(obj)) {
      properties[k] = inferSchema(v);
      required.push(k);
    }
    const node: SchemaNode = { type: "object", properties };
    if (required.length > 0) node.required = required;
    return node;
  }

  if (typeof value === "string")  return { type: "string" };
  if (typeof value === "boolean") return { type: "boolean" };
  if (typeof value === "number")  return Number.isInteger(value) ? { type: "integer" } : { type: "number" };

  return {};
}

function mergeSchemas(schemas: SchemaNode[]): SchemaNode {
  if (schemas.length === 1) return schemas[0];

  const types = new Set(schemas.map((s) => s.type as string));
  if (types.size > 1) return { oneOf: schemas };

  const type = schemas[0].type;

  if (type === "object") {
    const allKeys = new Set(
      schemas.flatMap((s) => Object.keys((s.properties as SchemaNode) ?? {}))
    );
    const properties: SchemaNode = {};
    // A key is required only if every schema includes it
    const requiredSets = schemas.map((s) => new Set<string>((s.required as string[]) ?? []));
    const required: string[] = [];

    for (const key of allKeys) {
      const keySchemas = schemas
        .filter((s) => (s.properties as SchemaNode)?.[key] !== undefined)
        .map((s) => (s.properties as SchemaNode)[key] as SchemaNode);
      properties[key] = mergeSchemas(keySchemas);
      if (requiredSets.every((set) => set.has(key))) required.push(key);
    }

    const merged: SchemaNode = { type: "object", properties };
    if (required.length > 0) merged.required = required;
    return merged;
  }

  return schemas[0];
}

export function generateJsonSchema(jsonString: string): string {
  const data = JSON.parse(jsonString);
  const schema = inferSchema(data);
  (schema as Record<string, unknown>).$schema = "http://json-schema.org/draft-07/schema#";
  // Move $schema to the top
  const { $schema, ...rest } = schema as { $schema: string } & SchemaNode;
  return JSON.stringify({ $schema, ...rest }, null, 2);
}
