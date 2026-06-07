type JsonValue = null | boolean | number | string | JsonValue[] | { [key: string]: JsonValue };

function escapeXml(str: string): string {
  return str
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

function sanitizeTag(name: string): string {
  let s = name.replace(/[^a-zA-Z0-9_.\-]/g, "_");
  if (/^[^a-zA-Z_]/.test(s)) s = "_" + s;
  return s || "_";
}

function toXmlNode(value: JsonValue, tag: string, depth: number): string {
  const indent = "  ".repeat(depth);

  if (value === null) return `${indent}<${tag} nil="true"/>`;
  if (typeof value === "boolean" || typeof value === "number") return `${indent}<${tag}>${value}</${tag}>`;
  if (typeof value === "string") return `${indent}<${tag}>${escapeXml(value)}</${tag}>`;

  if (Array.isArray(value)) {
    if (value.length === 0) return `${indent}<${tag}/>`;
    const children = value.map((item) => toXmlNode(item, "item", depth + 1)).join("\n");
    return `${indent}<${tag}>\n${children}\n${indent}</${tag}>`;
  }

  const entries = Object.entries(value);
  if (entries.length === 0) return `${indent}<${tag}/>`;
  const children = entries.map(([k, v]) => toXmlNode(v, sanitizeTag(k), depth + 1)).join("\n");
  return `${indent}<${tag}>\n${children}\n${indent}</${tag}>`;
}

export const jsonToXml = (json: string): string => {
  const parsed = JSON.parse(json) as JsonValue;
  return `<?xml version="1.0" encoding="UTF-8"?>\n${toXmlNode(parsed, "root", 0)}`;
};
