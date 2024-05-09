#!/usr/bin/env python
# coding: utf-8

# In[62]:


pip install pyperclip


# In[2]:


import tkinter as tk
from tkinter import ttk, scrolledtext, simpledialog, messagebox
import pyperclip

class Application(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("BW Formula Builder")
        self.geometry("1200x550")
        
        self.exclude_include_var = tk.StringVar(value="Include")
        self.custom_label_var = tk.StringVar(value="")
        self.field_var = tk.StringVar()
        self.value_vars = {}
        self.selected_values = {}
        self.generated = False
        self.saved_formulas = {}  # Dictionary to store saved formulas
        self.filtered_values = None  # Store filtered values for search

        self.load_fields_and_values()
        self.create_widgets()

        self.selected_indices = []
        
    def load_fields_and_values(self):
        # Load fields and values from the text file
        file_path = r"C:\Users\206791803\NBCUniversal\People Analytics Team - People Analytics\Team\Mike and Antoine\BW Formula Generator\bw formula gen.txt"
        with open(file_path, 'r', encoding='utf-8') as file:
            for line in file:
                line = line.strip()
                if ':' in line:
                    current_field, values = line.split(':', 1)
                    current_field = current_field.strip()
                    values = [value.strip() for value in values.split(',')]
                    self.value_vars[current_field] = {value: tk.BooleanVar(value=False) for value in values}

        # Load saved formulas from the saved_formulas.txt file
        saved_formulas_file_path = r"C:\Users\206791803\NBCUniversal\People Analytics Team - People Analytics\Team\Mike and Antoine\BW Formula Generator\saved_formulas.txt"
        with open(saved_formulas_file_path, 'r', encoding='utf-8') as file:
            formula_name = ''
            formula_content = ''
            for line in file:
                line = line.strip()
                if ':' in line:
                    if formula_name and formula_content:
                        self.saved_formulas[formula_name] = formula_content
                    formula_name, formula_content = line.split(": ", 1)
                    formula_name = formula_name.strip()
                    formula_content = formula_content.strip("{").strip("}")
                else:
                    formula_content += '\n' + line.strip()
            if formula_name and formula_content:
                self.saved_formulas[formula_name] = formula_content

    def load_saved_formula(self):
        selected_formula_name = self.saved_formulas_dropdown.get()
        if selected_formula_name:
            formula = self.saved_formulas.get(selected_formula_name)
            if formula:
                # Clear the output text widget
                self.output_text.delete(1.0, tk.END)
                # Insert each line of the formula
                for line in formula.split('\n'):
                    self.output_text.insert(tk.END, line + '\n')
                messagebox.showinfo("Success", f"Loaded formula '{selected_formula_name}'.")
                # Disable the Generate button
                self.generate_button.config(state=tk.DISABLED)  
                self.append_button.config(state=tk.NORMAL)
            else:
                messagebox.showinfo("Error", f"Formula '{selected_formula_name}' not found.")

    def filter_values_listbox(self, event=None):
        query = self.search_entry.get().lower()
        values = list(self.value_vars.get(self.field_var.get(), {}).keys())
        self.values_listbox.delete(0, tk.END)
        for idx, value in enumerate(values):
            if query in value.lower():
                self.values_listbox.insert(tk.END, value)
                # Check if the index is in the stored selection and select it
                if idx in self.selected_indices:
                    self.values_listbox.selection_set(tk.END)

    def create_widgets(self):
        
        input_frame = ttk.Frame(self)
        input_frame.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")

        ########################################## EXCLUDE/INCLUDE FRAME #########################################################
        
        exclude_include_frame = ttk.LabelFrame(input_frame, text="Label Your Criteria", padding=(10, 5))
        exclude_include_frame.grid(row=0, column=0, padx=5, pady=5, sticky="ew")

        exclude_radio = ttk.Radiobutton(exclude_include_frame, text="Exclude", variable=self.exclude_include_var, value="Exclude")
        exclude_radio.grid(row=0, column=0, padx=5, pady=5)

        include_radio = ttk.Radiobutton(exclude_include_frame, text="Include", variable=self.exclude_include_var, value="Include")
        include_radio.grid(row=0, column=1, padx=5, pady=5)
        
        custom_radio = ttk.Radiobutton(exclude_include_frame, text="Custom", variable=self.exclude_include_var, value="Custom")
        custom_radio.grid(row=0, column=2, padx=5, pady=5)
        
        custom_label_entry = ttk.Entry(exclude_include_frame, textvariable=self.custom_label_var)
        custom_label_entry.grid(row=0, column=3, padx=5, pady=5)

        field_frame = ttk.LabelFrame(input_frame, text="Choose a Field", padding=(10, 5))
        field_frame.grid(row=1, column=0, padx=5, pady=5, sticky="ew")

        field_combobox = ttk.Combobox(field_frame, textvariable=self.field_var, values=list(self.value_vars.keys()), state="readonly", width=30)
        field_combobox.grid(row=0, column=0, padx=5, pady=5)
        field_combobox.bind("<<ComboboxSelected>>", self.update_checkboxes)
        
        
       ########################################## VALUE FRAME #########################################################
    
        value_frame = ttk.LabelFrame(input_frame, text="Select Values", padding=(10, 5))
        value_frame.grid(row=2, column=0, padx=5, pady=5, sticky="ew")

        self.values_listbox = tk.Listbox(value_frame, selectmode=tk.MULTIPLE, width=35, height=8, font=('Arial', 10))
        self.values_listbox.grid(row=1, column=0, padx=5, pady=5, sticky="nsew")   

        scrollbar = ttk.Scrollbar(value_frame, orient=tk.VERTICAL, command=self.values_listbox.yview)
        scrollbar.grid(row=1, column=1, sticky="ns")
        self.values_listbox.config(yscrollcommand=scrollbar.set)

        self.generate_button = ttk.Button(value_frame, text="Generate", command=self.generate_output)
        self.generate_button.grid(row=3, column=0, padx=5, pady=5, sticky="ew")

        self.append_button = ttk.Button(value_frame, text="Append", command=self.append_output, state=tk.DISABLED)
        self.append_button.grid(row=3, column=1, padx=5, pady=5, sticky="ew")

        clear_button = ttk.Button(value_frame, text="Clear All", command=self.clear_selections)
        clear_button.grid(row=3, column=2, padx=5, pady=5, sticky="ew")
        
        scrollbar = ttk.Scrollbar(value_frame, orient=tk.VERTICAL, command=self.values_listbox.yview)
        scrollbar.grid(row=1, column=1, sticky="ns")
        self.values_listbox.config(yscrollcommand=scrollbar.set) 
        self.values_listbox.bind("<<ListboxSelect>>", self.update_selected_indices)
        
        search_frame = ttk.Frame(value_frame)
        search_frame.grid(row=0, column=0, columnspan=2, padx=5, pady=5, sticky="ew")

        self.search_entry = ttk.Entry(search_frame)
        self.search_entry.grid(row=0, column=0, padx=5, pady=5, sticky="ew")
        self.search_entry.bind("<KeyRelease>", self.filter_values_listbox)

        search_button = ttk.Button(search_frame, text="Search", command=self.filter_values_listbox)
        search_button.grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        
        clear_selection_button = ttk.Button(search_frame, text="Clear Selection", command=self.clear_selection_only)
        clear_selection_button.grid(row=0, column=2, padx=5, pady=5, sticky="ew")
        
        ########################################## INPUT FRAME #########################################################
        
        output_frame = ttk.LabelFrame(input_frame, text="Output Preview", padding=(10, 5))
        output_frame.grid(row=0, column=1, rowspan=3, padx=10, pady=5, sticky="nsew")

        self.output_text = scrolledtext.ScrolledText(output_frame, wrap=tk.WORD, width=70, height=20, font=('Arial', 10))
        self.output_text.grid(row=0, column=0, padx=5, pady=5, sticky="nsew")

        buttons_frame = ttk.Frame(input_frame)
        buttons_frame.grid(row=3, column=0, columnspan=2, padx=5, pady=5, sticky="ew")

        copy_button = ttk.Button(input_frame, text="Copy", command=self.copy_output)
        copy_button.grid(row=0, column=16, padx=5, pady=5, sticky="ew")
        
        # Dropdown menu to load saved formulas
        saved_formulas_frame = ttk.Frame(input_frame)
        saved_formulas_frame.grid(row=4, column=0, columnspan=2, padx=5, pady=5, sticky="ew")

        # Save Button
        save_button = ttk.Button(saved_formulas_frame, text="Save", command=self.save_formula)
        save_button.grid(row=0, column=2, padx=5, pady=5, sticky="ew")
        
        self.saved_formulas_var = tk.StringVar()
        self.saved_formulas_dropdown = ttk.Combobox(saved_formulas_frame, values=list(self.saved_formulas.keys()), width=30, state="readonly")
        self.saved_formulas_dropdown.grid(row=0, column=0, padx=5, pady=5, sticky="ew")

        load_saved_formula_button = ttk.Button(saved_formulas_frame, text="Load Saved Formula", command=self.load_saved_formula)
        load_saved_formula_button.grid(row=0, column=1, padx=5, pady=5, sticky="ew")

        # Rename Button
        rename_button = ttk.Button(saved_formulas_frame, text="Rename", command=self.rename_formula)
        rename_button.grid(row=0, column=4, padx=5, pady=5, sticky="ew")
        
        delete_button = ttk.Button(saved_formulas_frame, text="Delete", command=self.delete_formula)
        delete_button.grid(row=0, column=5, padx=5, pady=5, sticky="ew")
  
    def copy_output(self):
        output_content = self.output_text.get(1.0, tk.END).strip()
        if output_content:
            pyperclip.copy(output_content)
            messagebox.showinfo("Success", "Formula Copied to Clipboard")
        else:
            messagebox.showinfo("Error", "Formula is Empty")

    def update_checkboxes(self, event=None):
        field = self.field_var.get()
        values = list(self.value_vars.get(field, {}).keys())
        self.values_listbox.delete(0, tk.END)
        for value in values:
            self.values_listbox.insert(tk.END, value)

    def clear_selections(self):
        self.selected_values = {}
        self.values_listbox.selection_clear(0, tk.END)
        self.output_text.delete(1.0, tk.END)
        self.generated = False
        self.generate_button.config(state=tk.NORMAL)
        self.generate_button.config(text="Generate")
        self.append_button.config(state=tk.DISABLED)
        self.saved_formulas_dropdown.set('')
        self.search_entry.delete(0, tk.END)
        
         # Clear the filtered values in the list box
        self.update_checkboxes()

        # Clear selected values and indices
        self.values_listbox.selection_clear(0, tk.END)
        self.selected_values = {}
        self.selected_indices = []

    def update_selected_indices(self, event=None):
        # Update the stored selected indices whenever the selection changes
        self.selected_indices = self.values_listbox.curselection()
        
        
    def generate_output(self):
        field = self.field_var.get()
        values = list(self.value_vars.get(field, {}).keys())

        # Retrieve selected values from the filtered listbox
        selected_values = [self.values_listbox.get(idx) for idx in self.values_listbox.curselection()]
        if not selected_values:
            messagebox.showwarning("Warning", "Please select at least one value.")
            return

        self.selected_values[field] = selected_values

        output = ""
        first_line = True
        if field == "Latest Hire Date":
            simpledialog.askstring("Specify Date Parameters", "Enter Start Date:")
            simpledialog.askstring("Specify Date Parameters", "Enter End Date:")

        else:
            for field, values in self.selected_values.items():
                exclude_include = self.exclude_include_var.get()
                custom_label = self.custom_label_var.get()
                for index, value in enumerate(values):
                    if exclude_include == "Custom":
                        output += f'=If ([{field}]) = "{value}" Then "{custom_label}"\n'
                    else:
                        if index == 0 and (first_line or self.generated):
                                output += f'=If [{field}] = "{value}" Then "{exclude_include} - {value}"\n'
                                first_line = False
                                self.generated = True
                        else:
                            output += f'ElseIf [{field}] = "{value}" Then "{exclude_include} - {value}"\n'
            self.output_text.insert(tk.END, output)
            self.generate_button.config(state=tk.DISABLED)
            self.append_button.config(state=tk.NORMAL)
            # clear selections
            self.values_listbox.selection_clear(0, tk.END)
            self.selected_values = {}

    def append_output(self):
        field = self.field_var.get()
        values = list(self.value_vars.get(field, {}).keys())

        # Retrieve selected values from the filtered listbox
        selected_values = [self.values_listbox.get(idx) for idx in self.values_listbox.curselection()]
        if not selected_values:
            messagebox.showwarning("Warning", "Please select at least one value.")
            return

        self.selected_values[field] = selected_values

        output = ""
        first_line = True
        for field, values in self.selected_values.items():
            exclude_include = self.exclude_include_var.get()
            custom_label = self.custom_label_var.get()
            for value in values:
                if exclude_include == "Custom":
                    output += f'=ElseIf [{field}] = "{value}" Then "{custom_label}"\n'
                else:
                    output += f'ElseIf [{field}] = "{value}" Then "{exclude_include} - {value}"\n'
        self.output_text.insert(tk.END, output)
        # clear selections
        self.values_listbox.selection_clear(0, tk.END)
        self.selected_values = {}


    def clear_selection_only(self):
        self.values_listbox.selection_clear(0, tk.END)
        self.selected_values = {}
        
         # Clear text in the search bar
        self.search_entry.delete(0, tk.END)

        # Clear the filtered values in the list box
        self.update_checkboxes()

        # Clear selected values and indices
        self.values_listbox.selection_clear(0, tk.END)
        self.selected_values = {}
        self.selected_indices = []

    def save_formula(self):
        saved_formulas_file_path = r"C:\Users\206791803\NBCUniversal\People Analytics Team - People Analytics\Team\Mike and Antoine\BW Formula Generator\saved_formulas.txt"
        selected_formula_name = self.saved_formulas_dropdown.get()
        output_content = self.output_text.get(1.0, tk.END)
        if output_content.strip() == "":
            messagebox.showinfo("Error", "No formula to save.")
            return
        if selected_formula_name:
            self.saved_formulas[selected_formula_name] = output_content.strip()
            # Update the saved formulas file with the updated content
            with open(saved_formulas_file_path, 'w') as file:
                for name, content in self.saved_formulas.items():
                    file.write(f"{name}: {{{content}}}\n")
            messagebox.showinfo("Success", f"Formula '{selected_formula_name}' saved successfully.")
        else:
            formula_name = simpledialog.askstring("Enter Formula Name", "Please enter a name for the saved formula:")
            if formula_name:
                output_content = self.output_text.get(1.0, tk.END)
                with open(saved_formulas_file_path, 'a', encoding='utf-8') as file:
                    file.write(f"{formula_name}: {{{output_content.strip()}}}\n")
                self.saved_formulas[formula_name] = output_content.strip()
                self.saved_formulas_dropdown['values'] = list(self.saved_formulas.keys())
                messagebox.showinfo("Success", f"Formula '{formula_name}' saved successfully.")

    def rename_formula(self):
        selected_formula_name = self.saved_formulas_dropdown.get()
        if selected_formula_name:
            new_name = simpledialog.askstring("Rename Formula", "Enter a new name for the selected formula:")
            if new_name:
                self.saved_formulas[new_name] = self.saved_formulas.pop(selected_formula_name)
                # Update the dropdown menu options
                self.saved_formulas_dropdown['values'] = list(self.saved_formulas.keys())
                # Select the renamed formula in the dropdown
                self.saved_formulas_dropdown.set(new_name)
                messagebox.showinfo("Success", f"Formula '{selected_formula_name}' renamed to '{new_name}'.")     
                
    def delete_formula(self):
        selected_formula_name = self.saved_formulas_dropdown.get()
        if selected_formula_name:
            confirmed = messagebox.askyesno("Confirm Delete", f"Are you sure you want to delete the formula '{selected_formula_name}'?")
            if confirmed:
                del self.saved_formulas[selected_formula_name]
                self.saved_formulas_dropdown['values'] = list(self.saved_formulas.keys())
                self.saved_formulas_dropdown.set('')  # Reset the selection
                messagebox.showinfo("Success", f"Formula '{selected_formula_name}' deleted successfully.")

if __name__ == "__main__":
    app = Application()
    app.focus_force()
    app.mainloop()


# In[ ]:




