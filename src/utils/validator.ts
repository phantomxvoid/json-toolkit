export const validateJson = (json: string) => {
  try {
    JSON.parse(json);

    return {
      valid: true,
      error: "",
    };
  } catch (error) {
    return {
      valid: false,
      error: (error as Error).message,
    };
  }
};