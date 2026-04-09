from flask import Flask, render_template, request, jsonify
import xlwings as xw

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
                return jsonify({"status": "error", "message": f"Missing {r}"})

        area = float(data["area"])
        persons = float(data["persons"])
        co2 = float(data["co2"])
        co2_median = float(data["co2_median"])

        # -----------------------------
        # OPEN EXCEL
        # -----------------------------
        app_excel = xw.App(visible=False)
        app_excel.display_alerts = False
        app_excel.screen_updating = False

        wb = app_excel.books.open(FILE_PATH)
        ws = wb.sheets[0]

        # -----------------------------
        # WRITE INPUTS
        # -----------------------------
        ws.range("B2").value = area
        ws.range("B5").value = persons
        ws.range("F2").value = co2

        # -----------------------------
        # FORCE CALCULATION
        # -----------------------------
        wb.app.calculate()

        # -----------------------------
        # GOAL SEEK CORE LOGIC
        # G305 MUST EQUAL CO2 MEDIAN
        # BY CHANGING B7
        # -----------------------------
        success = ws.range("G305").api.GoalSeek(
            Goal=co2_median,
            ChangingCell=ws.range("B7").api
        )

        # -----------------------------
        # READ OUTPUTS
        # -----------------------------
        b7_value = ws.range("B7").value
        g305_value = ws.range("G305").value

        wb.save()

        if not success:
            return jsonify({
                "status": "error",
                "message": "Goal Seek did not converge",
                "B7": b7_value,
                "G305": g305_value
            })

        return jsonify({
            "status": "success",
            "B7": b7_value,
            "G305": g305_value,
            "target_CO2_median": co2_median
        })

    except Exception as e:
        return jsonify({
            "status": "error",
            "message": str(e)
        })

    finally:
        if wb:
            wb.close()
        if app_excel:
            app_excel.quit()


if __name__ == "__main__":
    app.run(debug=True)