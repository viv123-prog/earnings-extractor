import streamlit as st
import fitz  # PyMuPDF
from openai import OpenAI
import pandas as pd
import re
import json
from openpyxl import load_workbook
from openpyxl.styles import Font
import os

# Ensure no proxy interference
for proxy_var in ["http_proxy", "https_proxy", "HTTP_PROXY", "HTTPS_PROXY"]:
    os.environ.pop(proxy_var, None)

# Initialize OpenAI using global config
try:
    if "openai_api_key" not in st.secrets:
        st.error("❌ Missing API key. Add it to secrets.toml as 'openai_api_key = \"sk-...\"'")
        st.stop()

    # Create OpenAI client with the API key
    client = OpenAI(api_key=st.secrets["openai_api_key"])

    with st.spinner("🔌 Testing OpenAI connection..."):
        # Test with a simple call (Model list)
        client.models.list()

    st.success("✅ OpenAI connected successfully!")

except Exception as e:
    st.error(f"🚨 OpenAI initialization failed. Error: {str(e)}")
    st.stop()

# Define financial metrics and their Indian-specific aliases
financial_metrics = [
    # Profit & Loss Statement Metrics
    "Total Revenue", "Net Profit", "EBITDA", "EBIT", "Gross Profit", 
    "Operating Profit", "Profit Before Tax", "Profit After Tax",
    "Other Income", "Exceptional Items", "Tax Expense", "Finance Costs",
    "Depreciation", "Amortization", "Employee Benefits Expense",
    
    # Balance Sheet Metrics
    "Share Capital", "Reserves and Surplus", "Total Borrowings",
    "Current Liabilities", "Non-Current Liabilities", 
    "Current Assets", "Fixed Assets", "Investments",
    "Cash and Bank Balances", "Loans and Advances",
    "Trade Receivables", "Inventory", "Trade Payables",
    
    # Segment Reporting
    "Segment Revenue - Domestic", "Segment Revenue - Exports",
    "Segment Revenue - Business Unit 1", "Segment Revenue - Business Unit 2",
    "Segment Profit - Domestic", "Segment Profit - Exports",
    
    # Cash Flow Statement
    "Cash Flow from Operations", "Cash Flow from Investing",
    "Cash Flow from Financing", "Net Cash Flow",
    "Cash and Cash Equivalents at End of Period",
    
    # Key Ratios and Per Share Data
    "EPS (Basic)", "EPS (Diluted)", "Book Value per Share",
    "Dividend per Share", "Dividend Payout Ratio",
    "Return on Equity", "Return on Capital Employed",
    "Shares Outstanding"
]

metric_aliases = {
    # Profit & Loss
    "Total Revenue": ["Turnover", "Total Income", "Net Sales", "Revenue from Operations"],
    "Net Profit": ["Profit for the Period", "PAT", "Net Profit After Tax"],
    "EBITDA": ["Operating Profit Before Depreciation", "PBITDA"],
    "Gross Profit": ["Gross Margin", "Trading Profit"],
    "Finance Costs": ["Interest Expense", "Borrowing Costs"],
    
    # Balance Sheet
    "Reserves and Surplus": ["Retained Earnings", "General Reserve"],
    "Total Borrowings": ["Total Debt", "Secured Loans", "Unsecured Loans"],
    "Trade Receivables": ["Sundry Debtors", "Accounts Receivable"],
    "Trade Payables": ["Sundry Creditors", "Accounts Payable"],
    
    # Segment Reporting
    "Segment Revenue - Domestic": ["India Revenue", "Local Market Sales"],
    "Segment Revenue - Exports": ["Overseas Revenue", "Foreign Sales"],
    
    # Cash Flow
    "Cash Flow from Operations": ["Net Cash from Operating Activities"],
    "Cash and Cash Equivalents at End of Period": ["Closing Cash Balance"],
    
    # Ratios
    "Return on Equity": ["ROE", "Return on Net Worth"],
    "Return on Capital Employed": ["ROCE", "Return on Investment"]
}

def extract_text_from_pdf(pdf_file):
    """Extract text from PDF with Indian financial report handling"""
    doc = fitz.open(stream=pdf_file.read(), filetype="pdf")
    text = ""
    for page in doc:
        text += page.get_text()
    
    # Clean Indian-specific formatting
    text = re.sub(r"\(Rs\. in [Cc]rores?\)", "", text)
    text = re.sub(r"\(₹ in [Cc]rores?\)", "", text)
    text = re.sub(r"\(USD in [Mm]illions?\)", "", text)
    return text

def detect_quarters(text):
    """Detect Indian financial quarters"""
    quarters = []
    quarter_patterns = [
        r"(Q[1-4]\s?(?:FY|FY'|FY\s?)\d{2,4})",
        r"(Q[1-4]\s?\d{4}-\d{2})",
        r"(?:Quarter|Quarterly)\s(?:Ended|ending)\s(\w+\s\d{4})",
        r"(Q[1-4]\s?\d{4})",
        r"((?:January|February|March|April|May|June|July|August|September|October|November|December)\s\d{4})",
    ]
    for pattern in quarter_patterns:
        matches = re.finditer(pattern, text, re.IGNORECASE)
        for match in matches:
            quarter = match.group(1).strip()
            quarter = re.sub(r"FY\s?'?(\d{2})", r"FY20\1", quarter)
            if quarter not in quarters:
                quarters.append(quarter)
    return sorted(quarters, key=lambda x: (x[-4:], x[:2]))

def extract_quarter_data(full_text, quarter_section):
    """Extract financial data with Indian terminology support"""
    aliases_prompt = "\n".join([f"- {metric} (aliases: {', '.join(aliases)})" 
                              for metric, aliases in metric_aliases.items()])
    
    prompt = f"""Extract these financial metrics for {quarter_section} from an Indian company report:
Recognize alternative names and Indian terminology:

{aliases_prompt}

Return ONLY numerical values without symbols/units. Use "N/A" if missing.
Format as JSON with double quotes.

Example:
{{
    "Total Revenue": "1000000",
    "Net Profit": "150000"
}}

Text to analyze:
{full_text}
"""
    try:
        response = client.chat.completions.create(
            model="gpt-3.5-turbo",
            messages=[{"role": "user", "content": prompt}],
            temperature=0,
            response_format={"type": "json_object"}
        )
        content = response.choices[0].message.content
        return json.loads(content)
    except Exception as e:
        st.error(f"Error extracting data for {quarter_section}: {e}")
        return {metric: "N/A" for metric in financial_metrics}

def calculate_financial_ratios(df):
    """Calculate Indian-specific financial ratios"""
    st.subheader("🧮 Financial Ratios (Indian GAAP/Ind AS)")
    
    ratios = pd.DataFrame(index=[
        "Gross Margin (%)", "Operating Margin (%)", "Net Profit Margin (%)",
        "Return on Equity (%)", "Return on Capital Employed (%)",
        "Current Ratio", "Quick Ratio", "Debt-to-Equity",
        "Interest Coverage Ratio", "Book Value per Share (₹)",
        "Dividend Payout Ratio (%)"
    ])
    
    try:
        for quarter in df.columns:
            # Extract values with fallbacks
            rev = pd.to_numeric(df.loc["Total Revenue", quarter], errors='coerce') if "Total Revenue" in df.index else 0
            gp = pd.to_numeric(df.loc["Gross Profit", quarter], errors='coerce') if "Gross Profit" in df.index else 0
            op = pd.to_numeric(df.loc["Operating Profit", quarter], errors='coerce') if "Operating Profit" in df.index else 0
            np_val = pd.to_numeric(df.loc["Net Profit", quarter], errors='coerce') if "Net Profit" in df.index else 0
            te = pd.to_numeric(df.loc["Shareholder Equity", quarter], errors='coerce') if "Shareholder Equity" in df.index else 0
            borrowings = pd.to_numeric(df.loc["Total Borrowings", quarter], errors='coerce') if "Total Borrowings" in df.index else 0
            interest = pd.to_numeric(df.loc["Finance Costs", quarter], errors='coerce') if "Finance Costs" in df.index else 1
            shares = pd.to_numeric(df.loc["Shares Outstanding", quarter], errors='coerce') if "Shares Outstanding" in df.index else 1
            dividends = pd.to_numeric(df.loc["Dividend per Share", quarter], errors='coerce') if "Dividend per Share" in df.index else 0
            
            # Calculate ratios
            ratios.at["Gross Margin (%)", quarter] = f"{(gp/rev*100):.1f}%" if rev != 0 else "N/A"
            ratios.at["Operating Margin (%)", quarter] = f"{(op/rev*100):.1f}%" if rev != 0 else "N/A"
            ratios.at["Net Profit Margin (%)", quarter] = f"{(np_val/rev*100):.1f}%" if rev != 0 else "N/A"
            ratios.at["Return on Equity (%)", quarter] = f"{(np_val/te*100):.1f}%" if te != 0 else "N/A"
            ratios.at["Debt-to-Equity", quarter] = f"{(borrowings/te):.2f}" if te != 0 else "N/A"
            ratios.at["Interest Coverage Ratio", quarter] = f"{(op/interest):.2f}" if interest != 0 else "N/A"
            ratios.at["Book Value per Share (₹)", quarter] = f"₹{(te/shares):.2f}" if shares != 0 else "N/A"
            ratios.at["Dividend Payout Ratio (%)", quarter] = f"{(dividends/(np_val/shares)*100):.1f}%" if np_val != 0 and shares != 0 else "N/A"
        
        st.dataframe(ratios.style.highlight_null(color='lightyellow'))
        return ratios
    
    except Exception as e:
        st.error(f"Ratio calculation error: {str(e)}")
        return pd.DataFrame()

def generate_ai_insights(df, ratios):
    """Generate insights with Indian financial context"""
    st.subheader("💡 AI-Generated Insights")
    
    data_summary = f"""
Financial Data:
{df.to_string()}

Financial Ratios:
{ratios.to_string() if not ratios.empty else 'No ratios calculated'}
"""
    
    prompt = f"""Analyze this Indian company's financial data and provide key insights:
{data_summary}

Focus on:
- Revenue growth and profitability trends
- Liquidity and leverage position
- Segment performance (Domestic vs Exports)
- Key ratio analysis
- Any red flags or exceptional items

Format as bullet points with emojis:
- 📈 Insight 1
- ⚠️ Insight 2
- 💰 Insight 3"""
    
    try:
        with st.spinner("Generating insights..."):
            response = client.chat.completions.create(
                model="gpt-3.5-turbo",
                messages=[{"role": "user", "content": prompt}],
                temperature=0.5
            )
            insights = response.choices[0].message.content
            st.markdown(insights)
    
    except Exception as e:
        st.error(f"Couldn't generate insights: {str(e)}")

def main():
    st.title("📊 Indian Financial Data Extractor")
    
    uploaded_file = st.file_uploader("Upload Annual Report/Results PDF", type=["pdf"])
    
    if uploaded_file is not None:
        full_text = extract_text_from_pdf(uploaded_file)
        st.success("PDF text extracted successfully!")
        
        if st.checkbox("Show extracted text"):
            st.text_area("Full Text", full_text, height=300)
        
        all_quarters = detect_quarters(full_text)
        if not all_quarters:
            st.warning("No quarters detected. Using default quarters.")
            all_quarters = ["Q1 FY2023", "Q2 FY2023", "Q3 FY2023", "Q4 FY2023"]
        else:
            st.write(f"Detected reporting periods: {', '.join(all_quarters)}")
        
        selected_quarters = st.multiselect(
            "Select periods to analyze", 
            all_quarters,
            default=all_quarters[:min(4, len(all_quarters))]
        )
        
        if st.button("Extract Financial Data"):
            if not selected_quarters:
                st.warning("Please select at least one period")
                return
                
            with st.spinner("Extracting financial data..."):
                df = pd.DataFrame(index=financial_metrics, columns=selected_quarters)
                
                for quarter in selected_quarters:
                    quarter_data = extract_quarter_data(full_text, quarter)
                    for metric in financial_metrics:
                        df.loc[metric, quarter] = quarter_data.get(metric, "N/A")
                
                st.success("Data extraction complete!")
                st.dataframe(df)
                
                ratios = calculate_financial_ratios(df)
                if not ratios.empty:
                    generate_ai_insights(df, ratios)
                
                # Excel Export
                output_excel = "indian_financial_analysis.xlsx"
                
                with pd.ExcelWriter(output_excel, engine='openpyxl') as writer:
                    df.to_excel(writer, sheet_name="Financial Data")
                    if not ratios.empty:
                        ratios.to_excel(writer, sheet_name="Financial Ratios")
                    
                    # Formatting
                    workbook = writer.book
                    for sheetname in writer.sheets:
                        ws = writer.sheets[sheetname]
                        for col in ws.columns:
                            for cell in col:
                                if cell.row == 1:  # Header row
                                    cell.font = Font(bold=True)
                
                with open(output_excel, "rb") as f:
                    st.download_button(
                        label="Download Analysis (Excel)",
                        data=f,
                        file_name="indian_financial_analysis.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

if __name__ == "__main__":
    main()
