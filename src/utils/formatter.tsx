export const formatJson = (json: string): string => {
  const parsed = JSON.parse(json);
  return JSON.stringify(parsed, null, 2);
};

export const minifyJson = (json: string): string => {
  const parsed = JSON.parse(json);
  return JSON.stringify(parsed);
};