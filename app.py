"""
app.py — Flask backend for the ZERO backtest dashboard.

Endpoints
---------
GET  /                     -> static/index.html
GET  /api/variations       -> dropdown data (14 variations + modes)
POST /api/analyze          -> body: params JSON -> KPIs + chart series + heatmap
POST /api/export           -> body: params JSON -> ZIP (PDF + XLSX) download

Run:  python app.py   →  http://127.0.0.1:5000
"""

from __future__ import annotations

import io

from flask import Flask, jsonify, request, send_file, send_from_directory

from exporter import make_bundle, make_pdf, make_xlsx
from processor import MODES, VARIATIONS, Params, analyze, summarize_for_api

app = Flask(__name__, static_folder="static", static_url_path="/static")


@app.get("/")
def index():
    return send_from_directory("static", "index.html")


@app.get("/api/variations")
def variations():
    return jsonify({
        "variations": VARIATIONS,
        "modes": [{"id": "general", "label": "General Mode"},
                  {"id": "profitable", "label": "Only Profitable Hour + Profitable Weekday"}],
    })


def _params_from_request() -> Params:
    body = request.get_json(force=True, silent=True) or {}
    return Params(
        capital=body.get("capital", 100_000),
        variation=body.get("variation", VARIATIONS[0]["id"]),
        mode=body.get("mode", "general"),
        position_distribution=bool(body.get("position_distribution", True)),
        ma_filter=bool(body.get("ma_filter", False)),
        ma_period=body.get("ma_period", 50),
        ma_condition=body.get("ma_condition", "touch"),
    )


@app.post("/api/analyze")
def analyze_endpoint():
    try:
        p = _params_from_request()
        result = summarize_for_api(analyze(p))
        return jsonify(result)
    except FileNotFoundError as e:
        return jsonify({"error": str(e)}), 404
    except (ValueError, TypeError) as e:
        return jsonify({"error": str(e)}), 400
    except Exception as e:  # pragma: no cover
        app.logger.exception("analyze failed")
        return jsonify({"error": f"Internal error: {e}"}), 500


@app.post("/api/export")
def export_endpoint():
    try:
        p = _params_from_request()
        data, fname = make_bundle(analyze(p))
        return send_file(
            io.BytesIO(data),
            mimetype="application/zip",
            as_attachment=True,
            download_name=fname,
        )
    except FileNotFoundError as e:
        return jsonify({"error": str(e)}), 404
    except (ValueError, TypeError) as e:
        return jsonify({"error": str(e)}), 400
    except Exception as e:  # pragma: no cover
        app.logger.exception("export failed")
        return jsonify({"error": f"Internal error: {e}"}), 500


def _download(build, mimetype, name_fmt):
    try:
        p = _params_from_request()
        a = analyze(p)
        stamp = __import__("datetime").datetime.now().strftime("%Y%m%d_%H%M")
        fname = name_fmt.format(variation=p.variation, stamp=stamp)
        return send_file(io.BytesIO(build(a)), mimetype=mimetype,
                         as_attachment=True, download_name=fname)
    except FileNotFoundError as e:
        return jsonify({"error": str(e)}), 404
    except (ValueError, TypeError) as e:
        return jsonify({"error": str(e)}), 400
    except Exception as e:  # pragma: no cover
        app.logger.exception("export failed")
        return jsonify({"error": f"Internal error: {e}"}), 500


@app.post("/api/export/pdf")
def export_pdf():
    return _download(make_pdf, "application/pdf",
                     "ZERO_BacktestReport_{variation}_{stamp}.pdf")


@app.post("/api/export/xlsx")
def export_xlsx():
    return _download(make_xlsx,
                     "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                     "ZERO_TradeLog_{variation}_{stamp}.xlsx")


if __name__ == "__main__":
    print("ZERO Backtest Dashboard  →  http://127.0.0.1:5000")
    app.run(host="127.0.0.1", port=5000, debug=False)
