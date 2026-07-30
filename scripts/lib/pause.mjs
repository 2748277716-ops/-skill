export class PauseError extends Error {
  constructor(code, message, evidence = {}) {
    super(message);
    this.name = "PauseError";
    this.code = code;
    this.evidence = evidence;
  }

  toJSON() {
    return {
      status: "paused",
      code: this.code,
      message: this.message,
      evidence: this.evidence,
    };
  }
}

export function asPausedResult(error) {
  if (error instanceof PauseError) return error.toJSON();
  throw error;
}