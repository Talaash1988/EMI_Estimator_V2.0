# EMI_Estimator_V2.0

📊 Excel Loan Amortization UDF & EMI Estimator
🚀 Unlocking the Power of LAMBDA in Microsoft Excel
Welcome to my Loan Amortization Model, where I leverage Modern Excel techniques—including the LAMBDA function—to automate financial calculations efficiently. This repository contains: ✔ Loan Amortization UDF using LAMBDA, SEQUENCE, SCAN, LET, PMT, PPMT, IPMT. ✔ EMI Estimator Excel Workbook comparing the traditional method vs. LAMBDA-based modeling in separate sheets. ✔ A structured loan amortization schedule, including outstanding balances, principal, interest, and closing amounts.

✨ Features
✅ LAMBDA-driven Loan Amortization Function – Automates financial calculations with dynamic arrays. ✅ Comparison of Traditional vs. LAMBDA Modeling – See how modern Excel techniques enhance efficiency. ✅ Fully Customizable & Interactive – Adjust loan amounts, rates, and terms easily. ✅ Built for Microsoft 365 Users – Utilizes Excel’s newest functions for seamless automation.

📌 Usage Instructions
1️⃣ Download the Excel Workbook from this repository. 2️⃣ Open the sheet showcasing LAMBDA-based loan amortization vs. Traditional EMI estimation. 3️⃣ To use the Loan Amortization UDF in your own workbook:

Go to Formulas → Name Manager → New.

Copy the function below and assign it a name (e.g., LoanAmortizationUDF).

excel
=LAMBDA(Loan_amt,Months,Rate,PaymentDate,
LET(
header,{"Period","PaymentDate","Outstanding Amt.","Payment","Principal","Interest","Closing Balance"},
Period,SEQUENCE(Months),
pmtdate,IFERROR(EDATE(PaymentDate,SEQUENCE(Months)),""),
Payment,SEQUENCE(Months,1,PMT(Rate/12,Months,-Loan_amt),0),
Principal,PPMT(Rate/12,Period,Months,-Loan_amt),
Interest,IPMT(Rate/12,Period,Months,-Loan_amt),
Closing,SCAN(Loan_amt,Interest-Payment,SUM),
Opening,DROP(VSTACK(Loan_amt,Closing),-1),
Total,HSTACK("Total","","",SUM(Payment),"",SUM(Interest),""),
VSTACK(header,HSTACK(Period,pmtdate,Opening,Payment,Principal,Interest,Closing),Total)))

4️⃣ Apply the function in any cell:

=LoanAmortizationUDF(500000, 60, 0.06, DATE(2025,1,1))
This will generate an amortization schedule for a ₹500,000 loan over 60 months at a 6% interest rate, starting from January 1, 2025.

🚀 Contributing
Have an Excel challenge or an idea for improvement? I love problem-solving and exploring modern Excel techniques! 🔹 Feel free to submit issues or share complex problems, and I’ll try to tackle them. 🔹 Let’s optimize financial workflows using Microsoft 365 innovations!

📢 Connect with Me
💡 Follow me on LinkedIn for more insights into Excel, Financial Modeling, and Microsoft 365. 💬 Drop challenges in the comments—let’s solve them together!

⚡ License
This project is open-source and free to use for educational and financial modeling purposes.
