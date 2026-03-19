# asprova-platform

## 依存関係のインストール

```bash
py -m pip install -r requirements.txt
```

`apps/viewer` と `apps/bridge` 用のパッケージはルートの `requirements.txt` に統合しています。

## 起動

```bash
py run.py
```

デフォルトは **viewer を同一プロセスで起動**（`0.0.0.0:5000`, `debug=True`）。bridge だけ / 両方を別プロセスで動かす場合:

```bash
py run.py bridge
py run.py all
```
