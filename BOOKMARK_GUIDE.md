# アプリのブックマーク（ホーム画面追加）について

## 📱 推奨URL

ホーム画面に追加する際は、以下のURLを使用してください：

```
https://ユーザー名.github.io/ishiharabrain-cmd/orientation-support-app.html
```

**理由**: 
- インストーラー画面（index.html）をスキップして、直接アプリ画面が開きます
- 初回設定後は、このURLから開く方が便利です

---

## 🔄 2つのURL比較

### 1. index.html（インストーラー付き）
```
https://ユーザー名.github.io/ishiharabrain-cmd/
```

**用途**: 
- 初めての人に配布
- 「今すぐ始める」から設定画面へ誘導

**デメリット**:
- 2回目以降も「今すぐ始める」画面が出る

---

### 2. orientation-support-app.html（直接起動）⭐推奨
```
https://ユーザー名.github.io/ishiharabrain-cmd/orientation-support-app.html
```

**用途**:
- 普段使い
- すでに設定済みの人

**メリット**:
- 直接アプリ画面が開く
- すぐに使える

---

## 📋 使い分け

### 初めての人に案内する場合
```
https://ユーザー名.github.io/ishiharabrain-cmd/
```
↓
インストーラー → 設定 → アプリ

### すでに使っている人
```
https://ユーザー名.github.io/ishiharabrain-cmd/orientation-support-app.html
```
↓
直接アプリ起動

---

## 🔧 既存ユーザーの変更方法

すでにホーム画面に追加済みの場合：

### 方法1: 削除して再追加
1. ホーム画面のアイコンを長押し
2. 「Appを削除」
3. 新しいURL（orientation-support-app.html）で再度追加

**メリット**: 確実
**デメリット**: 手間がかかる
**注意**: 設定データは消えません

---

### 方法2: そのまま使う
- 設定は保存されているので、「今すぐ始める」を毎回スキップすればOK
- 特に問題なければそのままでも使えます

---

## 💡 推奨する配布方法

### QRコードを2種類用意

**QRコード1: 初回用**
```
https://ユーザー名.github.io/ishiharabrain-cmd/
```
「初めての方はこちら」

**QRコード2: リピーター用**
```
https://ユーザー名.github.io/ishiharabrain-cmd/orientation-support-app.html
```
「すでに設定済みの方はこちら」

---

## 📝 案内文の例

### 初めての方向け
```
【いつでもサポート アプリのご案内】

以下のリンクをタップしてください：
https://あなたのユーザー名.github.io/ishiharabrain-cmd/

1. 「今すぐ始める」をタップ
2. 基本情報を入力
3. ホーム画面に追加
```

### すでに使っている方向け
```
【いつでもサポート アプリ】

こちらから直接アクセスできます：
https://あなたのユーザー名.github.io/ishiharabrain-cmd/orientation-support-app.html

ホーム画面に追加すると便利です。
```

---

## ⚙️ manifest.jsonの設定

Version 1.1.1から、PWAのstart_urlを以下に変更しました：

```json
"start_url": "orientation-support-app.html"
```

**効果**:
- ホーム画面から起動すると直接アプリ画面が開く
- インストーラー画面をスキップ

---

**おすすめ: orientation-support-app.html をブックマークしてください！**
