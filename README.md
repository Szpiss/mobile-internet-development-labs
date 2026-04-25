# 移动互联网开发基础实验仓库

本仓库用于存放《移动互联网开发基础》课程的各次实验源码、实验报告和相关素材。

当前设备环境：

- macOS
- MacBook Pro M1 Pro
- Visual Studio Code
- Google Chrome

## 仓库说明

当前仓库已收录实验一和实验二相关内容，后续实验可以继续在本仓库中追加提交。

建议后续按实验编号整理目录，例如：

```text
lab1/
lab2/
lab3/
```

当前仓库暂时以“源码 + 实验文档 + 素材”方式整理，后续也可以继续按这个思路维护。

## 当前内容

- `index.html`
  实验一网页源码，实现 `1! + 2! + ... + N!` 的计算与显示。
- `实验二/`
  实验二网页源码目录，当前包含 `index.html`，用于演示 JavaScript 事件和表单页面设计。
- `实验文档/`
  统一存放实验说明与实验报告。
  当前包含实验一和实验二的实验说明与实验报告。
- `report_assets/`
  实验报告中使用的截图、流程图等素材。

## 实验一内容

实验一主题：

开发环境搭建并编写简单的 JavaScript 函数。

功能要求：

- 使用 `meta` 标签设置网页属性
- 设置网页背景色和文字颜色
- 使用 `rgb()` 函数设置颜色
- 使用文本标签组织页面内容
- 使用 JavaScript 计算 `1! + 2! + ... + N!`
- 显示中间计算过程和最终结果
- 支持清空输入和结果

## 实验二内容

实验二主题：

JavaScript 事件和表单页面设计。

功能要求：

- 点击超链接切换网页正文字号
- 演示三种事件绑定方式
- 演示窗口、鼠标和键盘事件状态
- 设计用户登录表单
- 对用户名和密码进行合法性校验
- 在提交与重置事件中给出交互反馈

## 运行方式

实验一直接在浏览器中打开：

```bash
open /Users/cuing/javascript/index.html
```

实验二直接在浏览器中打开：

```bash
open '/Users/cuing/javascript/实验二/index.html'
```

或者使用 VS Code 打开仓库目录后运行对应页面。

## GitHub 仓库

远程仓库地址：

[Szpiss/mobile-internet-development-labs](https://github.com/Szpiss/mobile-internet-development-labs)

## 后续维护建议

- 每次实验完成后及时提交源码和实验报告
- 截图素材统一放入对应实验目录或 `report_assets/`
- 临时文件、虚拟环境和缓存文件不要提交到仓库
