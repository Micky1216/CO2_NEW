from flask import Flask, render_template, request, jsonify
import xlwings as xw
import os
import traceback

app = Flask(__name__)

FILE_PATH = "data.xlsx"


@app.route("/")
def home():
    return render_template("index.html")


@app.route("/process", methods=["POST"])
def process():
    app_excel = None
    wb = None

    try:
        data = request.json

        # -----------------------------
        # INPUT VALIDATION
        # -----------------------------
        required = ["area", "persons", "co2", "co2_median"]
        for r in required:
            if r not in data or data[r] == "":
                return jsonify({
                    "status": "error",
                    "message": f"Missing {r}"
                })

        area = float(data["area"])
        persons = float(data["persons"])
        co2 = float(data["co2"])
        co2_median = float(data["co2_median"])

        # -----------------------------
        # OPEN EXCEL SAFELY
        # -----------------------------
        app_excel = xw.App(visible=False)
        app_excel.display_alerts = False
        app_excel.screen_updating = False

        wb = app_excel.books.open(os.path.abspath(FILE_PATH))
        ws = wb.sheets["Inquinanti_insieme"]

        # -----------------------------
        # WRITE INPUTS
        # -----------------------------
        ws.range("B2").value = area
        ws.range("B5").value = persons
        ws.range("F2").value = co2

        # -----------------------------
        # RECALCULATE
        # -----------------------------
        wb.app.calculate()

        # -----------------------------
        # GOAL SEEK
        # -----------------------------
        success = ws.range("G305").api.GoalSeek(
            Goal=co2_median,
            ChangingCell=ws.range("B7").api
        )

        # -----------------------------
        # READ RESULTS
        # -----------------------------
        b7_value = ws.range("B7").value
        g305_value = ws.range("G305").value

        wb.save()

        # -----------------------------
        # HANDLE FAILURE
        # -----------------------------
        if not success:
            return jsonify({
                "status": "error",
                "message": "Goal Seek did not converge",
                "B7": b7_value,
                "G305": g305_value
            })

        # -----------------------------
        # SUCCESS RESPONSE
        # -----------------------------
        return jsonify({
            "status": "success",
            "B7": round(b7_value, 4) if b7_value else b7_value,
            "G305": round(g305_value, 4) if g305_value else g305_value,
            "target_CO2_median": co2_median
        })

    except Exception as e:
        print("\n--- ERROR START ---")
        print(traceback.format_exc())
        print("--- ERROR END ---\n")

        return jsonify({
            "status": "error",
            "message": str(e)
        })

    finally:
        # -----------------------------
        # CLEANUP (SAFE)
        # -----------------------------
        try:
            if wb is not None:
                wb.close()
        except:
            pass

        try:
            if app_excel is not None:
                app_excel.quit()
        except:
            pass


if __name__ == "__main__":
    app.run(debug=True)
