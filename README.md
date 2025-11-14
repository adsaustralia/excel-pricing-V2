
# Excel SQM & Pricing Calculator (Streamlit)

This is a Streamlit app that:
- Uploads an Excel file containing material, size, and quantity data
- Detects materials and groups them into categories
- Lets you enter per-client pricing (AUD/m²) for each category
- Applies single-sided or double-sided pricing rules (via multipliers)
- Calculates SQM and prices per column and writes results back into the Excel file
- Shows previews in the browser and lets you download the updated workbook

## 🛠 Requirements

- Python 3.9+
- The packages listed in `requirements.txt`

Install them with:

```bash
pip install -r requirements.txt
```

## ▶️ Run the app locally

```bash
streamlit run app.py
```

Then open the URL shown in your terminal (usually `http://localhost:8501`).

## ⚙️ Usage

1. Upload your `.xlsx` Excel file.
2. Adjust the **Excel Structure Settings** if your rows/columns differ.
3. Choose **Single sided** or **Double sided**.
4. Review the detected material categories.
5. Enter the **base single-sided rate** (AUD/m²) for each category for this client.
6. Click **Process & Calculate**.
7. View per-sheet previews, combined totals, and download the updated Excel file.

## ☁️ Deploying to Streamlit Cloud

1. Push this folder (containing `app.py`, `requirements.txt`, and `README.md`) to a GitHub repo.
2. Go to [share.streamlit.io](https://share.streamlit.io) and sign in with GitHub.
3. Create a new app, select your repo, branch, and `app.py` as the entry point.
4. Deploy and share your Streamlit app URL with your team or clients.
