📊 Amazon Sales Analysis (Python)

A Python-based data analysis project that processes Amazon order raw data and generates clean summaries, insights, and visual charts.

This project is beginner-friendly and designed to demonstrate real-world data analysis + GitHub workflow.


---

📁 Project Structure

amazon-sales-report/
│
├── amazon_report.py        # Main Python analysis script
├── .gitignore              # Ignored files (env, cache, temp files)
├── README.md               # Project documentation


---

📌 Dataset Description

The raw Amazon order file contains the following important columns:

amazon-order-id

purchase-date

product-name

sku

item-status

quantity

item-price

item-tax

shipping-price

shipping-tax

ship-city

ship-state


> ⚠️ Raw Excel data is not uploaded to GitHub to maintain data privacy.




---

🧠 Analysis Performed

✅ Sheet 1: Cleaned Sales Data

Validated headers

Calculated total sales

Highlighted key columns using different colors


✅ Sheet 2: State-wise & Product-wise Sales

Sorted by State → Product → Sales

Charts included:

📊 Bar chart: State with highest product sales

🥧 Pie chart: Product-wise sales distribution


Charts use distinct colors for better readability


✅ Sheet 3: Cancelled & Returned Orders

Filtered orders where status = Cancelled / Returned

Separate summary for loss analysis



---

📊 Visualizations

The script automatically generates:

Bar charts (state-wise product sales)

Pie charts (product contribution)

Clean Excel formatting with highlighted columns


All charts use multiple attractive colors.


---

🛠️ Technologies Used

Python

Pandas

Matplotlib

OpenPyXL

Git & GitHub



---

▶️ How to Run This Project

1️⃣ Install dependencies

pip install pandas matplotlib openpyxl

2️⃣ Run the script

python amazon_report.py

3️⃣ Output

A formatted Excel report with multiple sheets

Charts embedded automatically



---

🌱 Learning Outcomes

Real-world Excel data cleaning

Python-based business analysis

Chart generation & formatting

GitHub project workflow



---

👤 Author

Shivangi Dhole
Data Analysis | Python | Excel | GitHub


---

⭐ If you find this useful, feel free to star the repository!
![dashboard](https://github.com/user-attachments/assets/b4a48dc5-5d34-434a-bcce-6887aaef30ed)

