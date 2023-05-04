import xlsxwriter
from flask import Flask, request, jsonify, send_file
from pymongo import MongoClient
from datetime import datetime, timedelta
import os

app = Flask(__name__)

# set up database connection
client = MongoClient(os.getenv("MONGO_URI"))
db = client["API"]
collection = db["revenue"]

# define data model
class Revenue:
    def __init__(self, category, type, date, amount):
        if category not in [
            "A. M. C. INCOME",
            "APPEALS FUND INCOME",
            "BOOKSHOP/SACRAMENTAL STORE INCOME",
            "BUILDING FUND INCOME / LAUNCHING",
            "CHAPEL / GROTTO BOXES INCOME",
            "CHURCH OFFERTORY INCOME/SUNDAYS",
            "CHURCH OFFERTORY INCOME / WEEKDAYS",
            "DONATIONS/GIFTS",
            "HARVEST INCOME",
            "HARVEST REDEMPTION (PLEDGE) INCOME",
            "INVESTMENT INCOME",
            "LOANS INCOME",
            "MANDATORY my_INCOME",
            "MASS INTENTIONS INCOME",
            "OUT-STATIONS INCOME",
            "PARISH CRUSADE/ADORATION/RETREAT/REVIVAL",
            "PARISH REGISTRATION INCOME",
            "RENTALS INCOME",
            "SACRAMENT OF CONFIRMATION INCOME",
            "SALES / BULLETIN, NEWSPAPER / FIXED ASSETS / MASS WINE  INCOME"
        ]:
            raise ValueError("Invalid revenue category")
        if type not in ["cash", "bank transfer"]:
            raise ValueError("Invalid revenue type. Must be one of 'cash' or 'bank transfer'")
        if not isinstance(amount, (int, float)) or amount <= 0:
            raise ValueError("Amount must be a positive number")
        self.category = category
        self.type = type
        self.date = date
        self.amount = amount

# create a new revenue
@app.route("/revenue", methods=["POST"])
def create_revenue():
    try:
        category = request.json["category"]
        type = request.json["type"]
        date = request.json["date"]
        amount = request.json["amount"]
        revenue = Revenue(category, type, date, amount).__dict__
        result = collection.insert_one(revenue)
        return jsonify({"message": "revenue created successfully.", "status":True})
    except KeyError as e:
        app.logger.debug(e)
        return jsonify({"error": f"Missing key in request body: {str(e)}", "status":False}), 400
    except ValueError as e:
        app.logger.debug(e)
        return jsonify({"error": str(e), "status":False}), 400
    except Exception as e:
        app.logger.debug(e)
        return jsonify({"error": "An error occurred while processing the request.", "status":False}), 500


class Expenses:
    def __init__(self, category, date, amount):
        if category not in ["appeals remittance", "assistance to religious", "assistance to smaller parishes", 
                            "auto maintenance", "auto purchase", "bank charges", "building expenses",
                            "catholic charities", "chancery approved levies", "chancery monthly returns",
                            "communications", "feeding", "furniture and equipment maintenance",
                            "furniture and equipment purchase", "harvest returns", "insurance",
                            "land purchases/improvement", "laundry/cleaning/fumigation", "legal", "loans repayment"]:
            raise ValueError("Invalid expense category")
        if not isinstance(amount, (int, float)) or amount <= 0:
            raise ValueError("Amount must be a positive number")
        self.category = category
        self.date = date
        self.amount = amount


# client = MongoClient(os.getenv("MONGO_URI"))
# db = client["API"]
my_collection = db["expenses"]

@app.route("/expenses", methods=["POST"])
def create_expense():
    try:
        category = request.json["category"]
        date = request.json["date"]
        amount = request.json["amount"]
        expense = Expenses(category, date, amount).__dict__
        result = my_collection.insert_one(expense)
        return jsonify({"message": "Expense created successfully.", "status":True})
    except KeyError as e:
        return jsonify({"error": f"Missing key in request body: {str(e)}", "status":False}), 400
    except ValueError as e:
        return jsonify({"error": str(e), "status":False}), 400
    except Exception as e:
        return jsonify({"error": "An error occurred while processing the request.", "status":False}), 500



@app.route("/reports", methods=["GET"])
def get_reports():
    try:
        start_date_str = request.args.get("start_date").strip()
        end_date_str = request.args.get("end_date").strip()
        start_date = datetime.strptime(start_date_str, "%Y-%m-%d")
        end_date = datetime.strptime(end_date_str, "%Y-%m-%d")

        query = {"date": {"$gte": start_date_str, "$lte": end_date_str}}
        results = collection.find(query)

        reports = []
        for result in results:
            report = {
                "category": result["category"],
                "type": result["type"],
                "date": result["date"],
                "amount": result["amount"]
            }
            reports.append(report)

        # Add one day if range is only one day
        if start_date == end_date:
            end_date = end_date + timedelta(days=1)

        # Create a new Excel workbook and add a worksheet
        filename = "reports.xlsx"
        workbook = xlsxwriter.Workbook(filename)
        worksheet = workbook.add_worksheet()

        # Write column headers
        headers = list(reports[0].keys())
        for i, header in enumerate(headers):
            worksheet.write(0, i, header)

        # Write data rows
        for i, report in enumerate(reports):
            for j, value in enumerate(report.values()):
                worksheet.write(i+1, j, value)

        # Close the workbook
        workbook.close()

        return send_file(filename, as_attachment=True)
    
    except ValueError:
        return({"error":"Invalid date format. Please use format: YYYY-MM-DD", "status":False}),400
    except KeyError:
        return({"error":"Invalid query. Please make sure that all fields are present in the collection.", "status":False}),500
    except IndexError:
        return({"error":"No data found for the given date range.", "status":False}),500
    except Exception as e:
        return({"error":str(e)}), 500
