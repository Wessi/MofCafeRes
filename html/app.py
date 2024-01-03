from flask import Flask, request, jsonify, render_template_string
import pyodbc
from datetime import datetime

app = Flask(__name__)

# Database connection (replace with your actual database path and credentials)
conn_str = (
    r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
    r'DBQ=C:\Users\Wessi\GitWS\MofCafeRes\html\cafedbhtml1.accdb;'
)
conn = pyodbc.connect(conn_str)
cursor = conn.cursor()

@app.route('/submit_transaction', methods=['POST'])
def submit_transaction():
    data = request.form
    # Extract data from form
    transaction_id = data['transactionId']
    #transaction_date = data['transactionDate']
    
    # Convert transaction_date to the proper datetime format expected by Access
    transaction_date = datetime.strptime(data['transactionDate'], '%Y-%m-%dT%H:%M').strftime('%Y-%m-%d %H:%M:%S')

    # Ensure that amount is a float or a decimal
    amount = float(data['amount']) if data['amount'] else 0.0
    
    transaction_type = data['transactionType']
    category = data['category']
    # amount = data['amount']
    payment_method = data['paymentMethod']
    description = data['description']
    employee_id = data['employeeId'] or None

    # SQL to insert data (modify the table and column names as per your database)
    sql = '''INSERT INTO Transactions 
             (TransactionID, TransactionDate, TransactionType, Category, Amount, PaymentMethod, Description, EmployeeID) 
             VALUES (?, ?, ?, ?, ?, ?, ?, ?)'''
    try:
        cursor.execute(sql, transaction_id, transaction_date, transaction_type, category, amount, payment_method, description, employee_id)
        conn.commit()
        return jsonify({"status": "success", "message": "Transaction recorded"})
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)})

# Add similar routes for other forms like 'submit_inventory', 'submit_daily_summary', 'submit_employee_info'

@app.route('/')
def index():
    # Render your HTML here
    return render_template_string(open(r'C:\Users\Wessi\GitWS\MofCafeRes\html\cafe.html').read())

if __name__ == '__main__':
    app.run(debug=True)
