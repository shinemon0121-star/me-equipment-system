# ME機器管理システム

医療機器の管理、貸出、点検、修理を統合的に管理するWebアプリケーションです。

## アクセス方法

このアプリケーはGitHub Pagesで公開されています：

**🔗 [ME機器管理システム - GitHub Pages](https://shinemon0121-star.github.io/me-equipment-system/)**

## 使用方法

1. 上記のリンクでアクセス
2. Google Apps ScriptのURLを設定
3. 各機能を利用開始

## 必要な設定

### Google Apps Script のデプロイURL

アプリはGoogle Apps ScriptをバックエンドAPIとして使用します。

1. [GAS_Code.gs](GAS_Code.gs)をコピー
2. [Google Apps Script](https://script.google.com)で新規プロジェクトを作成
3. コードを貼り付けて保存
4. **デプロイ** → **ウェブアプリ** として公開
5. 発行されたURLをアプリ内で設定

## 機能

- **機器マスター**: 医療機器の登録・管理
- **貸出履歴**: 機器の貸出・返却管理
- **点検記録**: 定期点検の記録
- **修理記録**: 修理依頼・修理内容の記録
- **スタッフ管理**: スタッフ情報の管理
- **部署管理**: 部署情報の管理
- **物品マスター**: 消耗品などの物品管理

## 技術スタック

- **フロントエンド**: HTML5 + Vanilla JavaScript
- **バックエンド**: Google Apps Script
- **データベース**: Google Sheets + IndexedDB
- **ホスティング**: GitHub Pages

## ブラウザ対応

- Chrome（推奨）
- Firefox
- Edge
- Safari

## 注意事項

- このアプリケーションはGoogle Sheetsと同期します
- オフライン対応（IndexedDB使用）
- 初回アクセス時にGAS URLの設定が必要です

---

**開発**: shinemon0121-star
**最終更新**: 2026年4月
