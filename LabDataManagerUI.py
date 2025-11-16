import pandas as pd
import tkinter as tk
import os
from tkinter import ttk, messagebox
from tkinter import MULTIPLE
from PIL import Image, ImageTk
from Functions import save_plots, export_to_excel, generate_pdf_report, import_sanitize  # Import functions from my_functions.py

# File to be updated
filename = 'lab_results_aug24.csv'
plots_dir = 'plots'  # Directory where the plot PNG files are stored

# Load and sanitize the data
df, grouped = import_sanitize(filename)

# Function to save new entry to CSV
def save_new_entry():
    bezeichnung = bezeichnung_var.get()
    wert = wert_var.get()
    datum = datum_var.get()
    if bezeichnung == "Neue Bezeichnung":
        bezeichnung = neue_bezeichnung_var.get()
        einheit = neue_einheit_var.get()
        if not bezeichnung or not einheit:
            messagebox.showerror("Input Error", "Please enter a new Bezeichnung and Einheit.")
            return
    else:
        einheit = einheit_label.cget("text")

    if not bezeichnung or not wert or not datum:
        messagebox.showerror("Input Error", "Please fill all fields.")
        return

    try:
        # Convert the value to a float
        wert = float(wert)
    except ValueError:
        messagebox.showerror("Input Error", "Please enter a valid numeric value for 'Wert'.")
        return

    try:
        # Validate the date format
        datum = pd.to_datetime(datum, format='%m/%d/%y').date()
    except ValueError:
        messagebox.showerror("Input Error", "Please enter the date in MM/DD/YY format.")
        return

    # Add a new row to the dataframe
    new_row = pd.DataFrame({
        'Bezeichnung': [bezeichnung],
        'Einheit': [einheit],
        'Wert': [wert],
        'Datum': [datum.strftime('%m/%d/%y')],
        'unterer Grenzwert': [None],
        'oberer Grenzwert': [None]
    })

    # Append the new row to the original dataframe
    global df, grouped
    df = pd.concat([df, new_row], ignore_index=True)

    # Save the updated dataframe back to the CSV file
    df.to_csv(filename, index=False)

    # Re-group the data after adding a new entry
    grouped = df.groupby('Bezeichnung')

    messagebox.showinfo("Success", "Entry saved successfully.")
    clear_fields()

# Function to update the unit when a Bezeichnung is selected
def update_unit(event):
    selected_bezeichnung = bezeichnung_var.get()
    if selected_bezeichnung == "Neue Bezeichnung":
        neue_bezeichnung_label.grid(column=0, row=3, padx=10, pady=10)
        neue_bezeichnung_entry.grid(column=1, row=3, padx=10, pady=10)
        neue_einheit_label.grid(column=0, row=4, padx=10, pady=10)
        neue_einheit_entry.grid(column=1, row=4, padx=10, pady=10)
        einheit_label.config(text="")
    else:
        neue_bezeichnung_label.grid_forget()
        neue_bezeichnung_entry.grid_forget()
        neue_einheit_label.grid_forget()
        neue_einheit_entry.grid_forget()
        if selected_bezeichnung:
            einheit = df.loc[df['Bezeichnung'] == selected_bezeichnung, 'Einheit'].iloc[0]
            einheit_label.config(text=einheit)
        else:
            einheit_label.config(text="")

# Function to clear input fields
def clear_fields():
    wert_var.set("")
    datum_var.set("")
    if bezeichnung_var.get() == "Neue Bezeichnung":
        neue_bezeichnung_var.set("")
        neue_einheit_var.set("")

# Function to load and display the selected graph
def show_graph(*args):
    selected_graph = graph_var.get()
    if selected_graph:
        graph_path = os.path.join(plots_dir, f"{selected_graph}.png")
        if os.path.exists(graph_path):
            img = Image.open(graph_path)
            img = img.resize((400, 300), Image.Resampling.LANCZOS)  # Resize the image for display
            img = ImageTk.PhotoImage(img)

            graph_label.config(image=img)
            graph_label.image = img
        else:
            messagebox.showerror("File Error", "Selected graph file does not exist.")
    else:
        messagebox.showerror("Selection Error", "Please select a graph to display.")

# Function to regenerate the plots
def regenerate_plots():
    global df, grouped
    # Re-generate the grouped object from the updated DataFrame
    df, grouped = import_sanitize(filename)

    # Generate the plots using the latest data
    save_plots(df, grouped)

    # Update the graph dropdown to reflect new plots
    update_graph_dropdown()
    messagebox.showinfo("Success", "Plots regenerated successfully.")


# Function to update the graph dropdown with the latest plots
def update_graph_dropdown():
    if os.path.exists(plots_dir):
        graph_files = [os.path.splitext(f)[0] for f in os.listdir(plots_dir) if f.endswith('.png')]
        graph_combobox['values'] = sorted(graph_files)

# Function to generate the Excel report and PDF
def generate_report():
    selected_indices = listbox.curselection()
    selected_values = [listbox.get(i) for i in selected_indices]
    if not selected_values:
        messagebox.showerror("Selection Error", "Please select at least one value for the report.")
        return

    # Generate Excel report
    export_to_excel(df, grouped, filename='lab_results_report.xlsx', selected_values=selected_values)

    # Generate PDF report
    generate_pdf_report(df, selected_values, filename='lab_results_report.pdf', plots_dir=plots_dir)

    messagebox.showinfo("Success", "Excel and PDF reports generated successfully.")

# Function to update Grenzwerte
def update_grenzwerte():
    selected_bezeichnung = grenzwert_bezeichnung_var.get()
    if not selected_bezeichnung:
        messagebox.showerror("Selection Error", "Please select a Bezeichnung.")
        return

    try:
        unterer_grenzwert = float(unterer_grenzwert_var.get()) if unterer_grenzwert_var.get() else None
        oberer_grenzwert = float(oberer_grenzwert_var.get()) if oberer_grenzwert_var.get() else None
    except ValueError:
        messagebox.showerror("Input Error", "Please enter valid numeric values for the Grenzwerte.")
        return

    # Update the grenzwerte in the dataframe
    global df, grouped
    df.loc[df['Bezeichnung'] == selected_bezeichnung, 'unterer Grenzwert'] = unterer_grenzwert
    df.loc[df['Bezeichnung'] == selected_bezeichnung, 'oberer Grenzwert'] = oberer_grenzwert

    # Format the Datum column before saving
    df['Datum'] = df['Datum'].dt.strftime('%m/%d/%y')

    # Save the updated dataframe back to the CSV file
    df.to_csv(filename, index=False)

    # Re-group the data after updating
    grouped = df.groupby('Bezeichnung')

    messagebox.showinfo("Success", "Grenzwerte updated successfully.")

    # Clear the fields after updating
    unterer_grenzwert_var.set("")
    oberer_grenzwert_var.set("")


# Function to populate the Grenzwerte fields based on the selected Bezeichnung
def populate_grenzwerte(event):
    selected_bezeichnung = grenzwert_bezeichnung_var.get()
    if selected_bezeichnung:
        existing_values = df[df['Bezeichnung'] == selected_bezeichnung].iloc[0]
        unterer_grenzwert_var.set(existing_values['unterer Grenzwert'] if pd.notna(existing_values['unterer Grenzwert']) else "")
        oberer_grenzwert_var.set(existing_values['oberer Grenzwert'] if pd.notna(existing_values['oberer Grenzwert']) else "")

# Create the main application window
root = tk.Tk()
root.title("Lab Results Manager")

# Section: Add New Entry
add_frame = ttk.LabelFrame(root, text="Add New Lab Result")
add_frame.grid(column=0, row=0, padx=10, pady=10, sticky="nsew")

# Bezeichnung dropdown
bezeichnung_label = ttk.Label(add_frame, text="Bezeichnung:")
bezeichnung_label.grid(column=0, row=0, padx=10, pady=10)
bezeichnung_var = tk.StringVar()
bezeichnung_combobox = ttk.Combobox(add_frame, textvariable=bezeichnung_var, state='readonly')
bezeichnung_combobox['values'] = ["Neue Bezeichnung"] + sorted(df['Bezeichnung'].unique())
bezeichnung_combobox.grid(column=1, row=0, padx=10, pady=10)
bezeichnung_combobox.bind('<<ComboboxSelected>>', update_unit)

# Display corresponding unit or allow input for new one
einheit_label = ttk.Label(add_frame, text="")
einheit_label.grid(column=2, row=0, padx=10, pady=10)

# Wert (Value) input
wert_label = ttk.Label(add_frame, text="Wert:")
wert_label.grid(column=0, row=1, padx=10, pady=10)
wert_var = tk.StringVar()
wert_entry = ttk.Entry(add_frame, textvariable=wert_var)
wert_entry.grid(column=1, row=1, padx=10, pady=10)

# Datum (Date) input
datum_label = ttk.Label(add_frame, text="Datum (MM/DD/YY):")
datum_label.grid(column=0, row=2, padx=10, pady=10)
datum_var = tk.StringVar()
datum_entry = ttk.Entry(add_frame, textvariable=datum_var)
datum_entry.grid(column=1, row=2, padx=10, pady=10)

# Input for new Bezeichnung and Einheit (hidden by default)
neue_bezeichnung_label = ttk.Label(add_frame, text="Neue Bezeichnung:")
neue_bezeichnung_var = tk.StringVar()
neue_bezeichnung_entry = ttk.Entry(add_frame, textvariable=neue_bezeichnung_var)
neue_einheit_label = ttk.Label(add_frame, text="Einheit:")
neue_einheit_var = tk.StringVar()
neue_einheit_entry = ttk.Entry(add_frame, textvariable=neue_einheit_var)

# Save button
save_button = ttk.Button(add_frame, text="Save", command=save_new_entry)
save_button.grid(column=0, row=5, columnspan=3, pady=20)

# Button to regenerate plots
regen_plots_button = ttk.Button(add_frame, text="Regenerate Plots", command=regenerate_plots)
regen_plots_button.grid(column=0, row=6, columnspan=3, pady=10)

# Section: Display Graphs
graph_frame = ttk.LabelFrame(root, text="Display Existing Graphs")
graph_frame.grid(column=1, row=0, padx=10, pady=10, sticky="nsew")

# Graph selection dropdown
graph_label_frame = ttk.Label(graph_frame, text="Select Graph:")
graph_label_frame.grid(column=0, row=0, padx=10, pady=10)
graph_var = tk.StringVar()
graph_combobox = ttk.Combobox(graph_frame, textvariable=graph_var, state='readonly')
graph_combobox.grid(column=1, row=0, padx=10, pady=10)
graph_combobox.bind('<<ComboboxSelected>>', show_graph)

# Populate the dropdown with existing PNG files in the 'plots' directory
update_graph_dropdown()

# Label to display the graph
graph_label = ttk.Label(graph_frame)
graph_label.grid(column=0, row=1, columnspan=3, padx=10, pady=10)

# Section: Generate Excel and PDF Report
report_frame = ttk.LabelFrame(root, text="Generate Reports")
report_frame.grid(column=0, row=1, columnspan=2, padx=10, pady=10, sticky="nsew")

# Listbox for selecting values to include in the reports
listbox_frame = ttk.Frame(report_frame)
listbox_frame.grid(column=0, row=0, padx=10, pady=10, sticky="nsew")

# Add a vertical scrollbar to the listbox
scrollbar = ttk.Scrollbar(listbox_frame, orient="vertical")
scrollbar.pack(side="right", fill="y")

listbox = tk.Listbox(listbox_frame, selectmode=MULTIPLE, yscrollcommand=scrollbar.set, height=10)
for value in sorted(df['Bezeichnung'].unique()):
    listbox.insert(tk.END, value)
listbox.pack(side="left", fill="both", expand=True)

scrollbar.config(command=listbox.yview)

# Button to generate the reports
generate_report_button = ttk.Button(report_frame, text="Generate Excel and PDF Report", command=generate_report)
generate_report_button.grid(column=0, row=1, padx=10, pady=10)

# Section: Edit Grenzwerte
grenzwert_frame = ttk.LabelFrame(root, text="Edit Grenzwerte")
grenzwert_frame.grid(column=1, row=1, padx=10, pady=10, sticky="nsew")

# Bezeichnung dropdown for Grenzwerte
grenzwert_bezeichnung_label = ttk.Label(grenzwert_frame, text="Bezeichnung:")
grenzwert_bezeichnung_label.grid(column=0, row=0, padx=10, pady=10)
grenzwert_bezeichnung_var = tk.StringVar()
grenzwert_bezeichnung_combobox = ttk.Combobox(grenzwert_frame, textvariable=grenzwert_bezeichnung_var, state='readonly')
grenzwert_bezeichnung_combobox['values'] = sorted(df['Bezeichnung'].unique())
grenzwert_bezeichnung_combobox.grid(column=1, row=0, padx=10, pady=10)
grenzwert_bezeichnung_combobox.bind('<<ComboboxSelected>>', populate_grenzwerte)

# Unterer Grenzwert input
unterer_grenzwert_label = ttk.Label(grenzwert_frame, text="Unterer Grenzwert:")
unterer_grenzwert_label.grid(column=0, row=1, padx=10, pady=10)
unterer_grenzwert_var = tk.StringVar()
unterer_grenzwert_entry = ttk.Entry(grenzwert_frame, textvariable=unterer_grenzwert_var)
unterer_grenzwert_entry.grid(column=1, row=1, padx=10, pady=10)

# Oberer Grenzwert input
oberer_grenzwert_label = ttk.Label(grenzwert_frame, text="Oberer Grenzwert:")
oberer_grenzwert_label.grid(column=0, row=2, padx=10, pady=10)
oberer_grenzwert_var = tk.StringVar()
oberer_grenzwert_entry = ttk.Entry(grenzwert_frame, textvariable=oberer_grenzwert_var)
oberer_grenzwert_entry.grid(column=1, row=2, padx=10, pady=10)

# Update Grenzwerte button
update_grenzwerte_button = ttk.Button(grenzwert_frame, text="Update Grenzwerte", command=update_grenzwerte)
update_grenzwerte_button.grid(column=0, row=3, columnspan=2, pady=10)

# Start the main event loop
root.mainloop()
