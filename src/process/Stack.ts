/**
 * 规则栈，用来保存规则和索引编号，同时避免重复
 */
export default class Stack {
  // 栈数据，用来保存规则和索引编号
  public readonly stack: Record<string, number> = {};

  // 当前索引
  private cursor: number;

  /**
   * 构造函数
   *
   * @param cursor - 索引从几开始
   */
  constructor(cursor: number = 1) {
    this.cursor = cursor;
  }

  /**
   * 尝试往栈里加入新规则并返回索引，如果规则已存在则返回直接返回相应索引
   *
   * @param rule - 规则
   * @return - 索引
   */
  public push(rule: string) {
    if (this.stack[rule]) {
      return this.stack[rule];
    }
    this.stack[rule] = this.cursor;
    this.cursor++;
    return this.stack[rule];
  }

  /**
   * 获取站内所有规则，按照索引升序排列
   *
   * @return 规则的数组
   */
  public getRules() {
    return Object.entries(this.stack)
      .sort((a, b) => a[1] - b[1])
      .map(([k, v]) => k.replace('$$id$$', v.toString()));
  }
}
