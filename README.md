# ImageLink for PowerPoint

ImageLink 是一个 Windows PowerPoint COM 加载项，用于把同一张设计图片在多个幻灯片中的实例建立关联。设计师更新其中一处时，插件可以同步替换当前演示文稿内的全部关联图片，同时分别保留各位置的尺寸、裁剪、旋转、翻转和普通动画。

## 主要功能

- 建立图片关联、加入已有关联、解除关联和查看关联位置。
- 一次更新同一关联组中的全部图片。
- 各关联位置独立保留图片框、裁剪焦点、旋转和翻转。
- 支持圆形、椭圆、圆角矩形、星形等 PowerPoint 内置“裁剪为形状”。
- 保留直接应用于图片的普通进入、强调、退出和动作路径动画。
- 同时提供功能区按钮和图片右键菜单。
- 成功通知默认关闭，可以在功能区中启用。
- 关联信息保存在 PPT 文件内部，不需要外部数据库或插件备份文件夹。

## 下载和安装

仓库中的 `installers` 目录提供两个安装程序：

- `ImageLink-User-Setup-0.2.0.exe`：仅为当前 Windows 用户安装，不需要管理员权限。
- `ImageLink-Setup-0.2.0.exe`：安装到当前计算机，需要管理员权限。

安装前请保存文件并完全退出 PowerPoint。运行对应安装程序，选择“安装 / 修复”，然后重新打开 PowerPoint。同一台电脑不要同时启用两个版本。

完整的分发压缩包位于 `packages` 目录。正式发布到 GitHub 时，更推荐把这两个 ZIP 上传到 GitHub Releases，并在代码仓库中只保留源码和文档。

## 使用方法

1. 选中一张普通嵌入图片，点击“建立关联”。
2. 使用复制粘贴把图片放到其他页面；也可以选中其他图片后点击“加入已有关联”。
3. 各位置可以独立调整大小、裁剪、旋转、翻转以及添加普通动画。
4. 选中关联组中的任意图片，点击“更新关联图片”，选择新的图片文件。
5. 检查各页构图并保存 PPT。

## 当前限制

- 仅支持普通嵌入图片，不支持组合内图片、图片占位符、外部链接图片或图形的图片填充。
- 支持 PowerPoint 内置“裁剪为形状”，暂不支持通过“编辑顶点”制作的自定义自由形状。
- 某些跨对象触发器、交互动作、复杂图片效果或对象关系会被安全检查阻止。
- 新旧图片比例差异较大时，插件会等比填满原图片框并尽量保留焦点，仍建议人工检查构图。

## 源代码目录

- `src/admin`：管理员版加载项、安装程序及测试源码。
- `src/user`：免管理员版加载项、安装程序及测试源码。
- `docs`：安装、测试和技术说明。

项目使用 C#、.NET Framework 和 PowerPoint COM 接口，不依赖第三方 NuGet 包。两个版本的功能逻辑相同，主要区别是安装目录和注册表范围。

## 构建

在 Windows PowerShell 中运行对应目录下的 `build.ps1`。脚本使用 Windows 自带的 .NET Framework C# 编译器，输出加载项 DLL、安装程序以及 32/64 位安装验证程序。

```powershell
powershell -ExecutionPolicy Bypass -File .\src\user\build.ps1 -OutputDirectory .\build\user
powershell -ExecutionPolicy Bypass -File .\src\admin\build.ps1 -OutputDirectory .\build\admin
```

构建和安装测试需要 Windows；实际功能测试需要桌面版 Microsoft PowerPoint。

## 版本

当前整理版本：0.2.0。

本仓库尚未选择开源许可证。公开到 GitHub 并不自动授予他人复制、修改或再发布源码的权利；如果希望接受外部使用或贡献，请在发布前选择合适的许可证。
