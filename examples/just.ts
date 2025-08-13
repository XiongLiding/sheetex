import { type JustData, justBuffer } from '../src/index.ts';
import { writeFileSync } from 'node:fs';

/**
 * 准备数据
 *
 * @remarks
 * 标题（caption）和表头（headers）都是可选的
 */
const data: JustData = {
  caption: '学生成绩',
  headers: ['序号', '姓名', '成绩'],
  data: [
    [1, '张三', 100],
    [2, '李四', 90],
    [3, '王五', 80],
  ],
};

const buffer = await justBuffer(data, [20, 20, 20]); // 列宽也可以不指定，但一般还是建议对实际需求中的数据量进行评估并设置
// 写入到指定位置
writeFileSync('./xlsx/just.xlsx', buffer);

// 快速获取 Blob 的方法
// const blob = await justBlob(data, [20, 20, 20]);

// 在浏览器中快速保存的方法
// await justSaveAs('just.xlsx', data, [20, 20, 20]);
