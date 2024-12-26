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
    """
    Fetches stock data for the entered tickers and selected metrics, then displays it in a table.

    This function also calculates the YTD performance for each stock and includes it in the table.
    """
    # Get stock tickers from input field and convert to uppercase
    stock_input = ticker_entry.get().upper()
    if not stock_input:
        messagebox.showwarning("Input Error", "Please enter stock tickers.")
        return

    # Split input into individual tickers
    tickers = [ticker.strip() for ticker in stock_input.split(',')]

    # Get selected metrics
    selected_metrics = [metric for metric,
                        var in metric_vars.items() if var.get()]
    if not selected_metrics:
        messagebox.showwarning(
            "Input Error", "Please select at least one metric.")
        return

    # Create a progress window to display while fetching data
    progress_window = create_progress_window()

    try:
        data = []

        # Add YTD performance row
        ytd_row = {'Metric': 'YTD Performance'}
        for ticker in tickers:
            update_progress(progress_window, f"Fetching data for {ticker}...")
            stock = yf.Ticker(ticker)
            hist = stock.history(start=f"{datetime.now().year}-01-01")
            if not hist.empty:
                start_price = hist.iloc[0]['Close']
                current_price = hist.iloc[-1]['Close']
                ytd_performance = (
                    (current_price - start_price) / start_price) * 100
                ytd_row[ticker] = f"{ytd_performance:.2f}%"
            else:
                ytd_row[ticker] = "N/A"
        data.append(ytd_row)

        # Fetch and format data for each selected metric
        for metric in selected_metrics:
            row = {'Metric': metric_display_names[metric]}
            for ticker in tickers:
                stock = yf.Ticker(ticker)
                info = stock.info
                value = info.get(metric, "N/A")
                if metric == "marketCap":
                    value = f"${
                        value / 1e9:.2f}B" if isinstance(value, (int, float)) else value
                elif metric in ["totalRevenue", "netIncome"]:
                    if metric == "netIncome":
                        value = info.get("netIncome") or info.get(
                            "netIncomeToCommon", "N/A")
                    value = f"${
                        value / 1e6:.2f}M" if isinstance(value, (int, float)) else value
                elif metric in ["profitMargins", "dividendYield"]:
                    value = f"{
                        value * 100:.2f}%" if isinstance(value, (int, float)) else value
                elif metric in ["trailingPE", "forwardPE", "priceToBook", "enterpriseToEbitda", "trailingEps", "beta"]:
                    value = f"{value:.2f}" if isinstance(
                        value, (int, float)) else value
                row[ticker] = value
            data.append(row)

        # Destroy progress window
        progress_window.destroy()

        # Create DataFrame from data
        df = pd.DataFrame(data)

        # Clear existing widgets
        for widget in frame.winfo_children():
            widget.destroy()

        # Recreate necessary widgets
        create_ui_elements()

        # Create table frame with weight
        table_frame = ttk.Frame(frame)
        table_frame.grid(row=4, column=0, columnspan=2,
                         sticky=(tk.W, tk.E, tk.N, tk.S), pady=(5))
        table_frame.grid_columnconfigure(0, weight=1)
        table_frame.grid_rowconfigure(0, weight=1)

        # Create Treeview
        global table
        table = ttk.Treeview(table_frame, columns=[
                             'Metric'] + tickers, show='headings', style="Custom.Treeview")

        # Set column headings and make them expand proportionally
        total_width = table_frame.winfo_width()
        metric_width = int(total_width * 0.3)  # 30% for metric column
        ticker_width = int((total_width - metric_width) /
                           len(tickers))  # Distribute remaining width
        table.heading('Metric', text='Metric')
        table.column('Metric', width=metric_width, minwidth=150, stretch=True)
        for ticker in tickers:
            table.heading(ticker, text=ticker)
            table.column(ticker, width=ticker_width,
                         minwidth=100, stretch=True)

        # Insert rows into the table
        for i, row in df.iterrows():
            table.insert('', 'end', values=list(row), tags=(
                'evenrow',) if i % 2 == 0 else ('oddrow',))

        # Configure row colors based on theme
        configure_table_style()

        # Add vertical scrollbar
        scrollbar = ttk.Scrollbar(
            table_frame, orient="vertical", command=table.yview)
        table.configure(yscrollcommand=scrollbar.set)

        # Grid layout for table
        table.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))

        # Enable copy and export buttons
        copy_button.config(state=tk.NORMAL)
        excel_button.config(state=tk.NORMAL)

    except Exception as e:
        progress_window.destroy()
        messagebox.showerror("Error", f"An error occurred: {str(e)}")


def create_progress_window():
    """
    Creates a progress window to display while fetching data.

    This window includes a label indicating that data is being fetched and a progress bar.
    """
    # Create a new Toplevel window for the progress display
    progress = tk.Toplevel(root)
    progress.title("Loading")  # Set the window title
    progress.geometry("300x100")  # Set the window size
    progress.transient(root)  # Make the window transient to the main window
    progress.grab_set()  # Grab focus to prevent interaction with the main window

    # Center the progress window
    progress.geometry("+%d+%d" % (root.winfo_x() + root.winfo_width() /
                      2 - 150, root.winfo_y() + root.winfo_height()/2 - 50))

    # Create a label to display the progress message
    ttk.Label(progress, text="Fetching data...").pack(pady=20)

    # Create an indeterminate progress bar
    progress_bar = ttk.Progressbar(progress, mode='indeterminate')
    progress_bar.pack(fill=tk.X, padx=20)  # Pack the progress bar horizontally
    progress_bar.start(10)  # Start the progress bar animation
    progress.update()  # Update the window to display the progress bar

    return progress  # Return the progress window object


def update_progress(progress_window, message):
    """
    Updates the progress window with a new message.

    Args:
        progress_window (tk.Toplevel): The progress window to update
        message (str): The new message to display
    """
    # Find the label widget in the progress window and update its text
    for widget in progress_window.winfo_children():
        if isinstance(widget, ttk.Label):
            widget.config(text=message)
    progress_window.update()  # Update the window to display the new message


def configure_table_style():
    """
    Configures the style of the table based on the current theme.

    This function updates the background and foreground colors of the table and its headings.
    """
    # Get the current theme
    current_theme = sv_ttk.get_theme()

    # Create a style object
    style = ttk.Style()

    # Configure the table style based on the theme
    if current_theme == "dark":
        style.configure("Custom.Treeview", background="#2a2a2a",
                        foreground="white", fieldbackground="#2a2a2a", rowheight=30)
        style.configure("Custom.Treeview.Heading", background="#3a3a3a",
                        foreground="white", font=('Helvetica', 10, 'bold'))
    else:
        style.configure("Custom.Treeview", background="white",
                        foreground="black", fieldbackground="white", rowheight=30)
        style.configure("Custom.Treeview.Heading", background="#f0f0f0",
                        foreground="black", font=('Helvetica', 10, 'bold'))

    # Map the selected row color based on the theme
    style.map('Custom.Treeview', background=[
              ('selected', '#4a90e2' if current_theme == "light" else '#2c5282')])


def create_ui_elements():
    """
    Creates the UI elements for the application.

    This function includes the header, input fields, metric selection, buttons, and table frame.
    """
    # Create header frame with weight
    header_frame = ttk.Frame(frame)
    header_frame.grid(row=0, column=0, columnspan=2,
                      sticky=(tk.W, tk.E), pady=(0, 20))
    header_frame.grid_columnconfigure(1, weight=1)

    # Create header label
    header_label = ttk.Label(
        header_frame, text="Stock Data Generator", font=('Helvetica', 16, 'bold'))
    header_label.grid(row=0, column=0, sticky=tk.W)

    # Create theme toggle button
    global theme_toggle_button
    theme_toggle_button = ttk.Button(
        header_frame, text="Switch Theme", command=toggle_theme)
    theme_toggle_button.grid(row=0, column=1, sticky=tk.E)

    # Create input frame with weight
    input_frame = ttk.LabelFrame(frame, text="Input", padding=(10, 5))
    input_frame.grid(row=1, column=0, columnspan=2,
                     sticky=(tk.W, tk.E), pady=(0, 10))
    input_frame.grid_columnconfigure(1, weight=1)

    # Create ticker label and entry
    ticker_label = ttk.Label(
        input_frame, text="Enter Stock Tickers (comma-separated):")
    ticker_label.grid(row=0, column=0, padx=(0, 10))
    global ticker_entry
    ticker_entry = ttk.Entry(input_frame)
    ticker_entry.grid(row=0, column=1, sticky=(tk.W, tk.E))

    # Create metrics frame with weight
    metrics_frame = ttk.LabelFrame(
        frame, text="Select Metrics", padding=(10, 5))
    metrics_frame.grid(row=2, column=0, columnspan=2,
                       sticky=(tk.W, tk.E), pady=(0, 10))
    metrics_frame.grid_columnconfigure(0, weight=1)

    # Create checkbox frame
    global checkbox_frame
    checkbox_frame = ttk.Frame(metrics_frame)
    checkbox_frame.grid(row=0, column=0, sticky=(tk.W, tk.E))

    # Calculate number of columns based on window width
    def update_checkbox_layout(event=None):
        """
        Updates the layout of the checkboxes based on the window width.

        This function recalculates the number of columns and rows for the checkboxes.
        """
        # Get the width of the metrics frame
        frame_width = metrics_frame.winfo_width()

        # Approximate width of each checkbox
        checkbox_width = 150

        # Calculate the number of columns and rows
        num_columns = max(2, frame_width // checkbox_width)
        num_rows = (len(available_metrics) + num_columns - 1) // num_columns

        # Remove existing checkboxes
        for widget in checkbox_frame.winfo_children():
            widget.grid_forget()

        # Recreate checkboxes in new layout
        for i, metric in enumerate(available_metrics):
            row = i % num_rows
            col = i // num_rows
            checkbox = ttk.Checkbutton(
                checkbox_frame, text=metric_display_names[metric], variable=metric_vars[metric])
            checkbox.grid(row=row, column=col, sticky=tk.W, padx=10, pady=2)

        # Configure grid weights for checkbox frame
        for i in range(num_columns):
            checkbox_frame.grid_columnconfigure(i, weight=1)

    # Bind the update function to the metrics frame
    metrics_frame.bind('<Configure>', update_checkbox_layout)

    # Create buttons frame
    global button_frame
    button_frame = ttk.Frame(frame)
    button_frame.grid(row=3, column=0, columnspan=2,
                      pady=(0, 10), sticky=(tk.W, tk.E))
    button_frame.grid_columnconfigure(3, weight=1)

    # Create generate button
    generate_button = ttk.Button(
        button_frame, text="Generate Table", command=generate_table)
    generate_button.grid(row=0, column=0, padx=5)

    # Create copy button
    global copy_button
    copy_button = ttk.Button(button_frame, text="Copy to Clipboard",
                             command=copy_to_clipboard, state=tk.DISABLED)
    copy_button.grid(row=0, column=1, padx=5)

    # Create Excel button
    global excel_button
    excel_button = ttk.Button(
        button_frame, text="Export to Excel", command=export_to_excel, state=tk.DISABLED)
    excel_button.grid(row=0, column=2, padx=5)


def copy_to_clipboard():
    """
    Copies the table data to the clipboard.

    This function retrieves the table data and formats it into a string that can be copied to the clipboard.
    """
    try:
        # Get table headers
        headers = []
        for col in table['columns']:
            headers.append(table.heading(col)['text'])

        # Get table data
        table_data = [headers]
        for item in table.get_children():
            row_data = table.item(item)['values']
            table_data.append(row_data)

        # Format data for clipboard
        clipboard_data = "\n".join(
            ["\t".join(map(str, row)) for row in table_data])

        # Clear and append data to clipboard
        root.clipboard_clear()
        root.clipboard_append(clipboard_data)
        root.update()

        # Show success message
        messagebox.showinfo("Success", "Data copied to clipboard!")

    except Exception as e:
        # Show error message
        messagebox.showerror(
            "Error", f"Could not copy data to clipboard: {str(e)}")


def export_to_excel():
    """
    Exports the table data to an Excel file.

    This function retrieves the table data and saves it to an Excel file named "stock_data.xlsx".
    """
    try:
        # Get table data
        table_data = []
        for item in table.get_children():
            row_data = table.item(item)['values']
            table_data.append(row_data)

        # Create DataFrame from table data
        df_export = pd.DataFrame(table_data[1:], columns=table_data[0])

        # Export to Excel
        df_export.to_excel("stock_data.xlsx", index=False)

        # Show success message
        messagebox.showinfo("Success", "Data exported to stock_data.xlsx!")

    except Exception as e:
        # Show error message
        messagebox.showerror(
            "Error", f"Could not export data to Excel: {str(e)}")


def toggle_theme():
    """
    Toggles the application theme between light and dark modes.

    This function updates the theme and configures the table style accordingly.
    """
    # Get current theme
    current_theme = sv_ttk.get_theme()

    # Toggle theme
    sv_ttk.set_theme("light" if current_theme == "dark" else "dark")

    # Configure table style
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

# Create UI elements
create_ui_elements()

# Set default theme to dark mode
sv_ttk.set_theme("dark")

# Start the main event loop
root.mainloop()
