function escapeCsv(value: string): string {
  if (value.includes(",") || value.includes('"') || value.includes("\n") || value.includes("\r")) {
    return `"${value.replace(/"/g, '""')}"`;
  }
  return value;
}

export function jsonToCsv(jsonString: string): string {
  const data = JSON.parse(jsonString);

  if (!Array.isArray(data)) {
    throw new Error("CSV export requires a JSON array at the root level.");
  }

  if (data.length === 0) {
    throw new Error("Array is empty — nothing to export.");
  }

  // Array of primitives
  const allPrimitive = data.every((v) => v === null || typeof v !== "object");
  if (allPrimitive) {
    return ["value", ...data.map((v) => escapeCsv(v == null ? "" : String(v)))].join("\n");
  }

  // Collect union of all keys (preserving first-seen order)
  const headerSet = new Set<string>();
  for (const row of data) {
    if (row !== null && typeof row === "object" && !Array.isArray(row)) {
      for (const key of Object.keys(row)) headerSet.add(key);
    }
  }

  const headers = Array.from(headerSet);
  if (headers.length === 0) {
    throw new Error("Could not determine columns from the array items.");
  }

  const lines: string[] = [headers.map(escapeCsv).join(",")];

  for (const row of data) {
    if (row === null || typeof row !== "object" || Array.isArray(row)) {
      lines.push(headers.map(() => "").join(","));
      continue;
    }
    const record = row as Record<string, unknown>;
    lines.push(
      headers.map((h) => {
        const val = record[h];
        if (val == null) return "";
        if (typeof val === "object") return escapeCsv(JSON.stringify(val));
        return escapeCsv(String(val));
      }).join(",")
    );
  }

  return lines.join("\n");
}
