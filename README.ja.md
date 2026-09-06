# ReoGrid — コミュニティエディション (V3)

[![NuGet](https://img.shields.io/nuget/v/unvell.ReoGrid.DLL.svg)](https://www.nuget.org/packages/unvell.ReoGrid.DLL/)
[![NuGet Downloads](https://img.shields.io/nuget/dt/unvell.ReoGrid.DLL.svg)](https://www.nuget.org/packages/unvell.ReoGrid.DLL/)
[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)
[![GitHub Stars](https://img.shields.io/github/stars/unvell/ReoGrid?style=social)](https://github.com/unvell/ReoGrid/stargazers)

**WinForms・WPF アプリケーションに Excel ライクなスプレッドシート機能を組み込める、高速・高機能な .NET コンポーネントです。**

> 🇺🇸 [English README](README.md)

---

## 🚀 ReoGrid V4 へのアップグレード

> **このリポジトリは ReoGrid V3（コミュニティエディション / MIT ライセンス）です。**
>
> **[ReoGrid V4](https://reogrid.net/jp/)** では、V3 をベースに大幅な機能強化が行われています。
> V4 は Professional 版・Enterprise 版の2エディションで提供（**機能は同一**）。サポート期間と導入規模でお選びください。

### V3 → V4 主な新機能

| 機能 | V3 コミュニティ | V4 |
|---|:---:|:---:|
| WinForms 対応 | ✅ | ✅ |
| WPF 対応 | ✅ | ✅（大幅強化） |
| Excel XLSX 入出力 | ✅ | ✅ |
| 数式・チャート | ✅ | ✅（数式を拡充） |
| **遅延ロード（100万行対応）** | — | ✅ |
| **複数行ヘッダー・セル結合** | — | ✅ |
| **データソース機能（IDataSource）** | — | ✅ |
| **条件付きスタイル** | — | ✅ |
| **カスタマイズ条件フィルター** | — | ✅ |
| **Excel 互換カスタム書式パターン** | — | ✅ |
| **セルロック 3段階制御** | — | ✅ |
| **WinAppDriver 自動テスト対応** | — | ✅ |
| **浮動小数点計算精度補正** | — | ✅ |
| 描画パフォーマンス | 高速 | さらに高速 |
| テクニカルサポート | — | ✅（付属） |
| ライセンス | MIT（無料） | 商用ライセンス |
| 価格 | 無料 | [価格一覧](https://reogrid.net/jp/prices) |

### エディション比較

| | Professional | Enterprise |
|---|:---:|:---:|
| 利用可能端末数 | 3台以下 | 無制限 |
| テクニカルサポート | 1ヶ月 | 3ヶ月 |
| 機能 | V4 フル機能 | V4 フル機能 |

👉 **[V4 の詳細・購入はこちら](https://reogrid.net/jp/)**

📧 ご質問・お見積りは [support@reogrid.net](mailto:support@reogrid.net) まで（日本語対応）

> **V3 から V4 への移行をご検討の方へ：** API の基本設計は互換性を保っており、スムーズに移行できます。移行に関するご相談も承っております。

---

## Web・クラウドの ReoGrid

ReoGrid はデスクトップだけではありません。同じ Excel ライクなエンジンが、いまブラウザでも動きます。

- **[ReoGrid Web](https://web.reogrid.net/jp)** — **React・Vue・バニラ JavaScript 向け**の高速なキャンバスベース スプレッドシートコンポーネント。ブラウザでの xlsx 入出力・数式・スタイル・条件付き書式。無料の [Lite ティア](https://web.reogrid.net/jp/pricing)あり。[ライブデモ](https://web.reogrid.net/jp/demos)や[無料オンライン XLSX ビューア](https://web.reogrid.net/jp/xlsx-viewer)をどうぞ。
- **[ReoGrid Studio](https://studio.reogrid.net)** — 公開 **MCP エンドポイント**を備えた、無料・ブラウザホスト型のスプレッドシート。自然言語でシートを作成、あるいは AI エージェントがスプレッドシートを公開し、共有可能なライブ URL を返します（インストール・ダウンロード不要）。

デスクトップツールを Web へ移行するなら → [ReoGrid .NET から ReoGrid Web へ](https://web.reogrid.net/jp/articles/reogrid-dotnet-to-web)。

---

## V3 コミュニティの機能

- **Excel (XLSX) 入出力**（OpenXML 経由）
- **リッチなセル書式**（フォント・サイズ・色・罫線・配置・折り返し・回転）
- **数式と関数**（SUM、COUNT、IF、VLOOKUP など）
- **チャート・画像・図形**
- **セル結合・固定・分割ビュー**
- **ソート・オートフィルター**
- **セルタイプとコントロール**（チェックボックス、ドロップダウン、ハイパーリンク、ボタンなど）
- **行・列のグループ化とアウトライン**
- **印刷とページ設定**
- **大規模ワークシートでも高速動作**
- **拡張可能なレンダリングとイベントモデル**

## 対応フレームワーク

| フレームワーク | ターゲット |
|---|---|
| .NET 8（Windows） | `net8.0-windows7.0` |
| .NET Framework 4.8 | `net48` |

## インストール

**.NET CLI**

```bash
dotnet add package unvell.ReoGrid.DLL
```

**Visual Studio パッケージマネージャーコンソール**

```powershell
Install-Package unvell.ReoGrid.DLL
```

## クイックスタート

### WinForms

```csharp
using unvell.ReoGrid;

var grid = new ReoGridControl { Dock = DockStyle.Fill };
this.Controls.Add(grid);

var sheet = grid.CurrentWorksheet;
sheet["A1"] = "こんにちは ReoGrid";
sheet.Cells["B1"].Data = DateTime.Now;

// XLSX の読み込み・保存
sheet.Load("input.xlsx");
sheet.Save("output.xlsx", FileFormat.Excel2007);
```

### WPF（XAML）

```xml
<Window
    x:Class="MyApp.MainWindow"
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    xmlns:rg="clr-namespace:unvell.ReoGrid;assembly=unvell.ReoGrid"
    Title="ReoGrid WPF" Height="450" Width="800">
  <Grid>
    <rg:ReoGridControl x:Name="grid"/>
  </Grid>
</Window>
```

```csharp
// コードビハインド
var sheet = grid.CurrentWorksheet;
sheet["A1"] = "こんにちは ReoGrid";
sheet["B1"] = 12345;
```

## スクリーンショット

Excel 読み込み
![Excel読み込み](https://raw.githubusercontent.com/unvell/ReoGrid/master/Snapshots/02.png)

チャート
![チャート](https://raw.githubusercontent.com/unvell/ReoGrid/master/Snapshots/276.png)

セルタイプとコントロール
![セルタイプとコントロール](https://raw.githubusercontent.com/unvell/ReoGrid/master/Snapshots/62.png)

行・列の固定
![行・列の固定](https://raw.githubusercontent.com/unvell/ReoGrid/master/Snapshots/08.png)

グループとアウトライン
![グループとアウトライン](https://raw.githubusercontent.com/unvell/ReoGrid/master/Snapshots/61.png)

印刷設定
![印刷設定](https://raw.githubusercontent.com/unvell/ReoGrid/master/Snapshots/01_2.png)

## ドキュメント

- ドキュメント・APIリファレンス・チュートリアル（英語）：**https://reogrid.net/document**
- 日本語サポートページ：**https://reogrid.net/jp**

## デモプロジェクト

`DemoWPF` ディレクトリに実行可能な WPF サンプルプロジェクトが含まれています。

## コントリビューション

V3 コミュニティへの貢献は歓迎します。詳しくは [CONTRIBUTING.md](CONTRIBUTING.md) をご覧ください。

> V3 の範囲を超える機能要望については、商用サポート付きで活発に開発が続く [ReoGrid V4 Enterprise](https://reogrid.net/jp/) をご検討ください。

## ライセンス

MIT ライセンス — Copyright (c) UNVELL Inc. 2012–2026, All rights reserved.

詳細は [LICENSE](LICENSE) をご確認ください。

---

## UNVELL について

ReoGrid は、.NET UI コンポーネント専門のソフトウェア会社 **[UNVELL Inc.](https://reogrid.net/jp/)** が開発・メンテナンスしています。

- 🌐 Webサイト：https://reogrid.net/jp
- 🧩 ReoGrid Web（JavaScript / React / Vue）：https://web.reogrid.net/jp
- ☁️ ReoGrid Studio（オンライン スプレッドシート＋MCP）：https://studio.reogrid.net
- 📧 サポート：support@reogrid.net
- 💬 日本語でのお問い合わせも承っております
