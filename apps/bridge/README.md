# ASPROVA BRIDGE for mcframe

Oracle（mcframe）からCSVを生成する簡易Webアプリ（Flask）です。

## 起動

PowerShell 例:

```powershell
cd C:\Users\lenovo\asprova-bridge
py -m pip install -r requirements.txt
py app.py
```

ブラウザで `http://localhost:5001/` を開きます。

## 使い方

1. 右上の **Connect mcframe** を押す
2. ポップアップで **ID / PASSWORD / SCHEMA** を入力して **Connect**
3. `Integrated Master` / `Item Table` を押してCSVを出力

※ Confirm で保存先フォルダを選べる機能は、ブラウザ（Chrome / Edge 等）によって対応状況が異なります。

