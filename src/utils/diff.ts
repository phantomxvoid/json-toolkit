type JsonValue =
  | null
  | boolean
  | number
  | string
  | JsonValue[]
  | { [key: string]: JsonValue };

export type DiffEntry = {
  type: "added" | "removed" | "changed";
  path: string;
  before?: JsonValue;
  after?: JsonValue;
};

export type DiffResult = {
  entries: DiffEntry[];
  summary: { added: number; removed: number; changed: number };
};

const isRecord = (value: JsonValue): value is { [key: string]: JsonValue } =>
  typeof value === "object" && value !== null && !Array.isArray(value);

const getType = (value: JsonValue) => {
  if (Array.isArray(value)) return "array";
  if (value === null) return "null";
  return typeof value;
};

const formatPath = (path: string, key: string | number) => {
  if (typeof key === "number") return `${path}[${key}]`;
  return /^[A-Za-z_$][\w$]*$/.test(key)
    ? `${path}.${key}`
    : `${path}[${JSON.stringify(key)}]`;
};

const compareValues = (
  before: JsonValue,
  after: JsonValue,
  path: string,
  diffs: DiffEntry[]
) => {
  if (Object.is(before, after)) return;

  const beforeType = getType(before);
  const afterType = getType(after);

  if (beforeType !== afterType) {
    diffs.push({ type: "changed", path, before, after });
    return;
  }

  if (Array.isArray(before) && Array.isArray(after)) {
    const maxLength = Math.max(before.length, after.length);
    for (let index = 0; index < maxLength; index += 1) {
      const nextPath = formatPath(path, index);
      if (index >= before.length) {
        diffs.push({ type: "added", path: nextPath, after: after[index] });
      } else if (index >= after.length) {
        diffs.push({ type: "removed", path: nextPath, before: before[index] });
      } else {
        compareValues(before[index], after[index], nextPath, diffs);
      }
    }
    return;
  }

  if (isRecord(before) && isRecord(after)) {
    const keys = Array.from(
      new Set([...Object.keys(before), ...Object.keys(after)])
    ).sort();
    keys.forEach((key) => {
      const nextPath = formatPath(path, key);
      if (!(key in before)) {
        diffs.push({ type: "added", path: nextPath, after: after[key] });
      } else if (!(key in after)) {
        diffs.push({ type: "removed", path: nextPath, before: before[key] });
      } else {
        compareValues(before[key], after[key], nextPath, diffs);
      }
    });
    return;
  }

  diffs.push({ type: "changed", path, before, after });
};

export const formatDiffAsText = (result: DiffResult): string => {
  if (result.entries.length === 0) return "No differences found.";
  const { added, removed, changed } = result.summary;
  const lines = [
    "JSON Diff Summary",
    `Added: ${added}`,
    `Removed: ${removed}`,
    `Changed: ${changed}`,
    "",
    ...result.entries.map((d) => {
      if (d.type === "added") return `+ ${d.path}: ${JSON.stringify(d.after)}`;
      if (d.type === "removed") return `- ${d.path}: ${JSON.stringify(d.before)}`;
      return `~ ${d.path}: ${JSON.stringify(d.before)} -> ${JSON.stringify(d.after)}`;
    }),
  ];
  return lines.join("\n");
};

export const diffJson = (originalJson: string, compareJson: string): DiffResult => {
  const original = JSON.parse(originalJson) as JsonValue;
  const compare = JSON.parse(compareJson) as JsonValue;
  const entries: DiffEntry[] = [];

  compareValues(original, compare, "$", entries);

  return {
    entries,
    summary: {
      added: entries.filter((d) => d.type === "added").length,
      removed: entries.filter((d) => d.type === "removed").length,
      changed: entries.filter((d) => d.type === "changed").length,
    },
  };
};
