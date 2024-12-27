import tkinter as tk
from tkinter import ttk
import yfinance as yf
import pandas as pd
import sv_ttk
from tkinter import messagebox
from datetime import datetime

# Define a dictionary to map metric keys to display names
metric_display_names = {
    'trailingPE': 'Trailing P/E',
    'forwardPE': 'Forward P/E',
    'beta': 'Beta',
    'marketCap': 'Market Cap',
    'priceToBook': 'Price to Book',
    'enterpriseToEbitda': 'EV/EBITDA',
    'trailingEps': 'EPS (TTM)',
    'totalRevenue': 'Revenue',
    'netIncome': 'Net Income',
    'profitMargins': 'Profit Margin',
    'dividendYield': 'Dividend Yield'
}

# Create list of available metrics
available_metrics = list(metric_display_names.keys())

def generate_table():
    stock_input = ticker_entry.get().upper()  # Convert to uppercase
    if not stock_input:
        messagebox.showwarning("Input Error", "Please enter stock tickers.")
        return
        
    tickers = [ticker.strip() for ticker in stock_input.split(',')]
    selected_metrics = [metric for metric, var in metric_vars.items() if var.get()]

    if not selected_metrics:
        messagebox.showwarning("Input Error", "Please select at least one metric.")
        return

    progress_window = create_progress_window()
    try:
        data = []
        for ticker in tickers:
            update_progress(progress_window, f"Fetching data for {ticker}...")
            row = {'Ticker': ticker}
            
            # Add YTD Performance
            stock = yf.Ticker(ticker)
            hist = stock.history(start=f"{datetime.now().year}-01-01")
            if not hist.empty:
                start_price = hist.iloc[0]['Close']
                current_price = hist.iloc[-1]['Close']
                ytd_performance = ((current_price - start_price) / start_price) * 100
                row['YTD Performance'] = f"{ytd_performance:.2f}%"
            else:
                row['YTD Performance'] = "N/A"

            # Add other metrics
            info = stock.info
            for metric in selected_metrics:
                value = info.get(metric, "N/A")
                if metric == "marketCap":
                    value = f"${value / 1e9:.2f}B" if isinstance(value, (int, float)) else value
                elif metric in ["totalRevenue", "netIncome"]:
                    if metric == "netIncome":
                        value = info.get("netIncome") or info.get("netIncomeToCommon", "N/A")
                    value = f"${value / 1e6:.2f}M" if isinstance(value, (int, float)) else value
                elif metric in ["profitMargins", "dividendYield"]:
                    value = f"{value * 100:.2f}%" if isinstance(value, (int, float)) else value
                elif metric in ["trailingPE", "forwardPE", "priceToBook", "enterpriseToEbitda", "trailingEps", "beta"]:
                    value = f"{value:.2f}" if isinstance(value, (int, float)) else value
                row[metric_display_names[metric]] = value
            data.append(row)
        
        progress_window.destroy()
        df = pd.DataFrame(data)
        
        # Clear existing widgets
        for widget in frame.winfo_children():
            widget.destroy()

        # Recreate necessary widgets
        create_ui_elements()

        # Create table frame with weight
        table_frame = ttk.Frame(frame)
        table_frame.grid(row=4, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(5))
        table_frame.grid_columnconfigure(0, weight=1)
        table_frame.grid_rowconfigure(0, weight=1)

        # Create Treeview
        global table
        metrics_columns = ['YTD Performance'] + [metric_display_names[m] for m in selected_metrics]
        table = ttk.Treeview(table_frame, columns=['Ticker'] + metrics_columns, show='headings', 
                            style="Custom.Treeview")

        # Set column headings and make them expand proportionally
        total_width = table_frame.winfo_width()
        ticker_width = int(total_width * 0.15)  # 15% for ticker column
        metric_width = int((total_width - ticker_width) / len(metrics_columns))  # Distribute remaining width

        table.heading('Ticker', text='Ticker')
        table.column('Ticker', width=ticker_width, minwidth=100, stretch=True)
        
        for metric in metrics_columns:
            table.heading(metric, text=metric)
            table.column(metric, width=metric_width, minwidth=100, stretch=True)

        # Insert rows into the table
        for i, row in df.iterrows():
            table.insert('', 'end', values=list(row), tags=('evenrow',) if i % 2 == 0 else ('oddrow',))

        # Configure row colors based on theme
        configure_table_style()

        # Add vertical scrollbar
        scrollbar = ttk.Scrollbar(table_frame, orient="vertical", command=table.yview)
        table.configure(yscrollcommand=scrollbar.set)

        # Grid layout for table
        table.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))

        copy_button.config(state=tk.NORMAL)
        excel_button.config(state=tk.NORMAL)

    except Exception as e:
        progress_window.destroy()
        messagebox.showerror("Error", f"An error occurred: {str(e)}")

def create_progress_window():
    progress = tk.Toplevel(root)
    progress.title("Loading")
    progress.geometry("300x100")
    progress.transient(root)
    progress.grab_set()
    
    # Center the progress window
    progress.geometry("+%d+%d" % (root.winfo_x() + root.winfo_width()/2 - 150,
                                 root.winfo_y() + root.winfo_height()/2 - 50))
    
    ttk.Label(progress, text="Fetching data...").pack(pady=20)
    progress_bar = ttk.Progressbar(progress, mode='indeterminate')
    progress_bar.pack(fill=tk.X, padx=20)
    progress_bar.start(10)
    
    progress.update()
    return progress

def update_progress(progress_window, message):
    for widget in progress_window.winfo_children():
        if isinstance(widget, ttk.Label):
            widget.config(text=message)
    progress_window.update()

def configure_table_style():
    current_theme = sv_ttk.get_theme()
    style = ttk.Style()
    
    if current_theme == "dark":
        style.configure("Custom.Treeview", background="#2a2a2a", foreground="white", 
                       fieldbackground="#2a2a2a", rowheight=30)
        style.configure("Custom.Treeview.Heading", background="#3a3a3a", foreground="white",
                       font=('Helvetica', 10, 'bold'))
    else:
        style.configure("Custom.Treeview", background="white", foreground="black", 
                       fieldbackground="white", rowheight=30)
        style.configure("Custom.Treeview.Heading", background="#f0f0f0", foreground="black",
                       font=('Helvetica', 10, 'bold'))
    
    style.map('Custom.Treeview', 
              background=[('selected', '#4a90e2' if current_theme == "light" else '#2c5282')])

def create_ui_elements():
    global theme_toggle_button, ticker_entry, checkbox_frame, button_frame, copy_button, excel_button, table

    # Header frame with weight
    header_frame = ttk.Frame(frame)
    header_frame.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 20))
    header_frame.grid_columnconfigure(1, weight=1)
    
    header_label = ttk.Label(header_frame, text="Stock Data Generator", 
                            font=('Helvetica', 16, 'bold'))
    header_label.grid(row=0, column=0, sticky=tk.W)
    
    theme_toggle_button = ttk.Button(header_frame, text="Switch Theme", command=toggle_theme)
    theme_toggle_button.grid(row=0, column=1, sticky=tk.E)

    # Input frame with weight
    input_frame = ttk.LabelFrame(frame, text="Input", padding=(10, 5))
    input_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
    input_frame.grid_columnconfigure(1, weight=1)
    
    ticker_label = ttk.Label(input_frame, text="Enter Stock Tickers (comma-separated):")
    ticker_label.grid(row=0, column=0, padx=(0, 10))
    
    ticker_entry = ttk.Entry(input_frame)
    ticker_entry.grid(row=0, column=1, sticky=(tk.W, tk.E))

    # Metrics frame with weight
    metrics_frame = ttk.LabelFrame(frame, text="Select Metrics", padding=(10, 5))
    metrics_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
    metrics_frame.grid_columnconfigure(0, weight=1)

    checkbox_frame = ttk.Frame(metrics_frame)
    checkbox_frame.grid(row=0, column=0, sticky=(tk.W, tk.E))

    # Calculate number of columns based on window width
    def update_checkbox_layout(event=None):
        frame_width = metrics_frame.winfo_width()
        checkbox_width = 150  # Approximate width of each checkbox
        num_columns = max(2, frame_width // checkbox_width)
        num_rows = (len(available_metrics) + num_columns - 1) // num_columns

        # Remove existing checkboxes
        for widget in checkbox_frame.winfo_children():
            widget.grid_forget()

        # Recreate checkboxes in new layout
        for i, metric in enumerate(available_metrics):
            row = i % num_rows
            col = i // num_rows
            checkbox = ttk.Checkbutton(checkbox_frame, text=metric_display_names[metric],
                                     variable=metric_vars[metric])
            checkbox.grid(row=row, column=col, sticky=tk.W, padx=10, pady=2)

        # Configure grid weights for checkbox frame
        for i in range(num_columns):
            checkbox_frame.grid_columnconfigure(i, weight=1)

    metrics_frame.bind('<Configure>', update_checkbox_layout)

    # Buttons frame
    button_frame = ttk.Frame(frame)
    button_frame.grid(row=3, column=0, columnspan=2, pady=(0, 10), sticky=(tk.W, tk.E))
    button_frame.grid_columnconfigure(3, weight=1)

    generate_button = ttk.Button(button_frame, text="Generate Table", command=generate_table)
    generate_button.grid(row=0, column=0, padx=5)
    
    copy_button = ttk.Button(button_frame, text="Copy to Clipboard", 
                            command=copy_to_clipboard, state=tk.DISABLED)
    copy_button.grid(row=0, column=1, padx=5)
    
    excel_button = ttk.Button(button_frame, text="Export to Excel", 
                             command=export_to_excel, state=tk.DISABLED)
    excel_button.grid(row=0, column=2, padx=5)

def copy_to_clipboard():
    try:
        headers = ['Ticker']  # Start with Ticker column
        for col in table['columns'][1:]:  # Skip the Ticker column since we already added it
            headers.append(table.heading(col)['text'])
        
        table_data = [headers]
        for item in table.get_children():
            row_data = table.item(item)['values']
            table_data.append(row_data)

        clipboard_data = "\n".join(["\t".join(map(str, row)) for row in table_data])
        
        root.clipboard_clear()
        root.clipboard_append(clipboard_data)
        root.update()
        
        messagebox.showinfo("Success", "Data copied to clipboard!")
        
    except Exception as e:
        messagebox.showerror("Error", f"Could not copy data to clipboard: {str(e)}")

def export_to_excel():
    try:
        headers = ['Ticker']  # Start with Ticker column
        for col in table['columns'][1:]:  # Skip the Ticker column since we already added it
            headers.append(table.heading(col)['text'])
        
        table_data = []
        for item in table.get_children():
            row_data = table.item(item)['values']
            table_data.append(row_data)

        df_export = pd.DataFrame(table_data, columns=headers)
        df_export.to_excel("stock_data.xlsx", index=False)
        messagebox.showinfo("Success", "Data exported to stock_data.xlsx!")
        
    except Exception as e:
        messagebox.showerror("Error", f"Could not export data to Excel: {str(e)}")

def toggle_theme():
    current_theme = sv_ttk.get_theme()
    sv_ttk.set_theme("light" if current_theme == "dark" else "dark")
    configure_table_style()

# Main Window Setup
root = tk.Tk()
root.title("Stock Data Generator")
root.geometry("1200x800")  # Larger initial size
root.minsize(800, 600)  # Set minimum window size

# Configure main frame to expand
frame = ttk.Frame(root, padding="20")
frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
frame.grid_columnconfigure(0, weight=1)
frame.grid_rowconfigure(4, weight=1)  # Make table row expandable

# Configure root window grid
root.grid_rowconfigure(0, weight=1)
root.grid_columnconfigure(0, weight=1)

# Initialize metric variables
metric_vars = {metric: tk.IntVar(value=0) for metric in available_metrics}

create_ui_elements()

sv_ttk.set_theme("dark")  # Default theme is dark mode

root.mainloop()