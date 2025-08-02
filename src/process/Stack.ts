export default class Stack {
  public readonly stack: Record<string, number> = {};
  private cursor: number;

  constructor(cursor: number = 1) {
    this.cursor = cursor;
  }

  public push(rule: string) {
    if (this.stack[rule]) {
      return this.stack[rule];
    }
    this.stack[rule] = this.cursor;
    this.cursor++;
    return this.stack[rule];
  }

  public getRules() {
    return Object.entries(this.stack)
      .sort((a, b) => a[1] - b[1])
      .map(([k, v]) => k.replace('{{id}}', v.toString()));
  }
}
