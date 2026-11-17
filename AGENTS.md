# CLAUDE.md

Behavioral guidelines to reduce common LLM coding mistakes. Merge with project-specific instructions as needed.

**Tradeoff:** These guidelines bias toward caution over speed. For trivial tasks, use judgment.

## 1. Think Before Coding

**Don't assume. Don't hide confusion. Surface tradeoffs.**

Before implementing:

- State your assumptions explicitly. If uncertain, ask.
- If multiple interpretations exist, present them - don't pick silently.
- If a simpler approach exists, say so. Push back when warranted.
- If something is unclear, stop. Name what's confusing. Ask.

## 2. Simplicity First

**Minimum code that solves the problem. Nothing speculative.**

- No features beyond what was asked.
- No abstractions for single-use code.
- No "flexibility" or "configurability" that wasn't requested.
- No error handling for impossible scenarios.
- If you write 200 lines and it could be 50, rewrite it.

Ask yourself: "Would a senior engineer say this is overcomplicated?" If yes, simplify.

## 3. Surgical Changes

**Touch only what you must. Clean up only your own mess.**

When editing existing code:

- Don't "improve" adjacent code, comments, or formatting.
- Don't refactor things that aren't broken.
- Match existing style, even if you'd do it differently.
- If you notice unrelated dead code, mention it - don't delete it.

When your changes create orphans:

- Remove imports/variables/functions that YOUR changes made unused.
- Don't remove pre-existing dead code unless asked.

The test: Every changed line should trace directly to the user's request.

## 4. Goal-Driven Execution

**Define success criteria. Loop until verified.**

Transform tasks into verifiable goals:

- "Add validation" → "Write tests for invalid inputs, then make them pass"
- "Fix the bug" → "Write a test that reproduces it, then make it pass"
- "Refactor X" → "Ensure tests pass before and after"

For multi-step tasks, state a brief plan:

```
1. [Step] → verify: [check]
2. [Step] → verify: [check]
3. [Step] → verify: [check]
```

Strong success criteria let you loop independently. Weak criteria ("make it work") require constant clarification.

## 5. 必须执行的项目技术规则

- 开发、修复、维护、重构、依赖或打包调整前，必须阅读 [rules 开发维护入口](rules/README.md) 和 [表格技术选型规范](rules/表格文件读写与修改技术选型规范.md)，按输入输出契约和操作编号选择技术。
- 必须检查本次功能及其直接调用链的读写效率；发现重复读写、调用开销或能力缺口时，主动提出包含代码/计时依据、候选技术、兼容范围和验收方法的具体调整方案。未测量不得宣称提速。
- 已授权的优化完成适配和验证；超出范围的跨功能迁移先提交可审阅方案。遵循规则中的默认选择、同契约用户备选和失败回退要求。
- 本仓库已停止开发，NoComPyTools 已移除；后续开发转至 https://github.com/Azi-lzb/Rust-Tools 。仅在用户明确要求时维护此历史版本，不重新创建子项目。
- 母项目保持 laptop 基线自动路由：Windows 自动路由功能中 `.xls` 走 COM，其他格式走 Python；COM 专用功能按其入口执行。
