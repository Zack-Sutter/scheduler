import copy
import datetime
import os
import random
import subprocess
import sys
import tkinter as tk
from tkinter import messagebox, ttk
from tksheet import Sheet

import numpy as np
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from openpyxl.formatting.rule import CellIsRule
from openpyxl.utils import get_column_letter

CELL_COLORS = {
    'Trike': '#9999FF',
    'Gallery': '#5eb91e',
    'Back': '#f4b183',
    'Front': '#8e86ae',
    'Float': '#D3D3D3',
    'ENCA': '#90E4C1',
    'Head': '#ffd221',
    'MOT': '#ffd221',
    'DIAR': '#b50934',
    'Lunch': '#7d7f7c',
    'Security': '#00b1d2',
    'Tickets': '#f6f9d4',
    'FLOAT': '#D3D3D3',
    'Badges': '#f6f9d4',
    'Project': '#f83dda',
    'GRGR': '#cf6498',
    'Manager': '#00008B',
    'STST': '#bf94e4',
    'MP': '#FFA500',
    'Museum Project': '3FFA500',
    'Training': '#FFD580',
    'Camp': '#c7ea46',
    'Retail': '#ffccd4',
    '': '#D3D3D3',
    None: '#D3D3D3'
}

primary_button_color = "#EDD863"
primary_button_hover_color = "#E1D591"
secondary_button_color = "#6A2E35"
secondary_button_hover_color = "#78454C"
#redo_button_color = "#E0ACD5"
#redo_hover_color = "#E6C7DF"
#generic_button_color = "#5C6B73"
#generic_hover_color = "#687278"

# RES_FILE_NAME = f'{'' if sys.platform == 'win32' else '/tmp/'}momath schedule {datetime.date.today()}.xlsx'

if sys.platform == 'win32':
    RES_FILE_NAME = f'momath_schedule_{datetime.date.today()}.xlsx'
else:
    RES_FILE_NAME = f'/tmp/momath_schedule_{datetime.date.today()}.xlsx'


class ScheduleApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title('MoMath Automatic Scheduler')
        self.geometry('1500x800')
        self.configure(background='lightblue')

        self.nonstandard_shifts = [] # renamed alr_ns
        self.action_history_stack = []
        self.action_redo_stack = []
        self.df = pd.DataFrame()
        self.sheet_frame = None
        self.create_widgets()

    def create_widgets(self):
        """Initialize all app widgets."""
        self.create_worker_inputs()
        self.create_radio_buttons()
        self.create_action_buttons()
        self.create_notes_box()

    def create_worker_inputs(self):
        """Initialize the worker entry boxes."""
        # Paid workers input
        tk.Label(self, text='Enter the name of FT/ROOT workers sparated by a comma (Ex: Daniel, Lou, Olivia, Sydney)').pack(anchor='w', pady=10, padx=10)
        self.paid_workers_entry = tk.Entry(self, width=65)
        #self.paid_workers_entry.insert(0, get_ft()) # get_ft() not added yet.
        self.paid_workers_entry.insert(0, "placeholder")
        self.paid_workers_entry.pack(anchor='w', pady=5, padx=10)

        # Volunteers input
        tk.Label(self, text='Enter the name of volunteers sparated by a comma (Ex: DKyle, Aliya, Alex)').pack(anchor='w', pady=5, padx=10)
        self.volunteers_entry = tk.Entry(self, width=65)
        self.volunteers_entry.pack(anchor='w', pady=5, padx=10)

    def create_radio_buttons(self):
        """Initialize radio buttons."""
        self.radio0 = tk.IntVar(value=1)
        tk.Label(self, text='Early or late lunches today?').pack(anchor='w', pady=5, padx=10)
        tk.Radiobutton(self, text='Early', variable=self.radio0, value=0).pack(anchor='w', pady=5, padx=10)
        tk.Radiobutton(self, text='Late', variable=self.radio0, value=1).pack(anchor='w', pady=5, padx=10)

        self.radio1 = tk.IntVar()
        tk.Label(self, text='Will FT/ROOT work tickets?').pack(anchor='w', pady=5, padx=10)
        tk.Radiobutton(self, text='No', variable=self.radio1, value=0).pack(anchor='w', pady=5, padx=10)
        tk.Radiobutton(self, text='Yes', variable=self.radio1, value=1).pack(anchor='w', pady=5, padx=10)

        self.radio2 = tk.IntVar(value=1)
        tk.Label(self, text='Will volunteers work tickets?').pack(anchor='w', pady=5, padx=10)
        tk.Radiobutton(self, text='No', variable=self.radio2, value=0).pack(anchor='w', pady=5, padx=10)
        tk.Radiobutton(self, text='Yes', variable=self.radio2, value=1).pack(anchor='w', pady=5, padx=10)

        self.radio4 = tk.IntVar()
        tk.Label(self, text='Is today a *Thursday* Freeplay?').pack(anchor='w', pady=5, padx=10)
        tk.Radiobutton(self, text='No', variable=self.radio4, value=0).pack(anchor='w', pady=5, padx=10)
        tk.Radiobutton(self, text='Yes', variable=self.radio4, value=1).pack(anchor='w', pady=5, padx=10)

    def create_action_buttons(self):
        """Initialize main buttons."""
        create_schedule_button = tk.Button(self, text="Create Schedule", command=self.create_schedule, background=primary_button_color)
        create_schedule_button.pack(anchor='w', pady=20, padx=10)
        create_schedule_button.bind('<Enter>', lambda e: create_schedule_button.configure(background=primary_button_hover_color))
        create_schedule_button.bind('<Leave>', lambda e: create_schedule_button.configure(background=primary_button_color))

        open_schedule_button = tk.Button(self, text='Open Schedule in Excel', command=self.open_excel, height=2, width=20, background=primary_button_color)
        open_schedule_button.place(relx=0.993, rely=0.89, anchor='e')
        open_schedule_button.bind('<Enter>', lambda e: open_schedule_button.configure(background=primary_button_hover_color))
        open_schedule_button.bind('<Leave>', lambda e: open_schedule_button.configure(background=primary_button_color))

        close_button = tk.Button(self, text='Close MAS', command=self.close, height=2, width=10, background=primary_button_color)
        close_button.place(relx=0.993, rely=0.95, anchor='e')
        close_button.bind('<Enter>', lambda e: close_button.configure(background=primary_button_hover_color))
        close_button.bind('<Leave>', lambda e: close_button.configure(background=primary_button_color))

    def create_notes_box(self):
        """Initialize Notes textbox."""
        self.notes_text_box = tk.Text(self, wrap='word', width=60, height=15)
        self.notes_text_box.place(relx=0.993, rely=0.70, anchor='e')
        self.load_notes()

    def destroy_sheet(self) -> None:
        """Destroys sheets."""
        for child in self.winfo_children():
            if isinstance(child, Sheet):
                child.destroy()

    def update_sheet(self):
        """Update the display sheet with the information in the dataframe."""
        self.sheet_frame.update_sheet()
        return
    
    def destroy_frame(window: tk.Tk) -> None:
        """Destroys frame."""
        for child in window.winfo_children():
            if isinstance(child, tk.Frame):
                child.destroy()

    def get_sheet_selection(self) -> dict:
        """Gets the selection of the sheet."""

        if not self.sheet_frame.sheet.get_all_selection_boxes():
            print("none selected")
            return
        selection = self.sheet_frame.sheet.get_all_selection_boxes()[0] # get first selection only
        time_start = self.df.iloc[selection.from_r].name
        time_end = self.df.iloc[selection.upto_r-1].name

        workers = self.sheet_frame.sheet.headers()[selection.from_c:selection.upto_c]

        return {"workers": workers, "time_start": time_start, "time_end": time_end}

    def add_nonstandard_shift(self, selection):
        """Handle adding nonstandard shifts to the dataframe and the display."""
        self.save_state()
        self.df.loc[selection["time_start"]:selection["time_end"], selection["workers"]] = selection["shift"]
        self.nonstandard_shifts.append(selection)
        self.update_sheet()

    def add_standard_shift(self, shift):
        """Handle adding standard shifts to the dataframe and the display."""
        if not self.sheet_frame.sheet.get_all_selection_boxes():
            workers = self.paid_workers + self.volunteers
        else:
            selection = self.sheet_frame.sheet.get_all_selection_boxes()[0] # get first selection only
            workers = self.sheet_frame.sheet.headers()[selection.from_c:selection.upto_c]

        failed_time_slots = []
        random.shuffle(workers)


        for i in range(len(self.df.index)):
            curr_row = self.df.index[i]
            workers_with_nan = set(self.df.columns[self.df.iloc[i].isna()].tolist()) & set(workers)
            if workers_with_nan:
                worker_to_assign = None
                for worker in workers:
                    if worker in workers_with_nan:
                        worker_to_assign = worker
                        break
                workers.remove(worker_to_assign)
                workers.append(worker_to_assign)
                self.df.at[curr_row, worker_to_assign] = shift
            if not self.df.loc[curr_row].isin([shift]).any():
                failed_time_slots.append(curr_row)
        if failed_time_slots and shift:
            msg = ''
            for index,slot in enumerate(failed_time_slots):
                if index > 0: 
                    msg += ', '
                msg += slot
            messagebox.showwarning('Warning', f'Failed to place {shift} at:\n'+ msg)

        self.update_sheet()
     

    def save_state(self):
        """Called when an action is done that wants to be undo-able."""
        self.action_history_stack.append(self.df.copy())
        self.action_redo_stack = []
 
    def undo(self):
        if self.action_history_stack:
            last_state = self.action_history_stack.pop()
            current_state = self.df.copy()
            self.action_redo_stack.append(current_state)
            self.df = last_state
            self.update_sheet()
        else:
            print("nothing left to undo.")  
    
    def redo(self):
        if self.action_redo_stack:
            next_state = self.action_redo_stack.pop()
            current_state = self.df.copy()
            self.action_history_stack.append(current_state)
            self.df = next_state
            self.update_sheet()
        else:
            print("nothing left to redo.")

    def load_notes(self):
        try:
            with open('daily_notes.txt') as file:
                content = file.read()
                self.notes_text_box.delete('1.0', tk.END)
                self.notes_text_box.insert(tk.END, content)
        except FileNotFoundError:
            self.notes_text_box.delete('1.0', tk.END)
            self.notes_text_box.insert(tk.END, 'Error: File not found.\nPlease make a file named daily_notes.txt and\nplace it in the same folder as schedule.py')
        except Exception as e:
            self.notes_text_box.delete('1.0', tk.END)
            self.notes_text_box.insert(tk.END, f'An error occurred: {e}')

    def save_notes(self):
        file_content = self.notes_text_box.get('1.0', tk.END)
        with open('daily_notes.txt', 'w') as file:
            file.write(file_content)

    def create_schedule(self):
        if self.sheet_frame: # clear previous UI element if button is clicked more than once.
            self.destroy_frame()

        self.paid_workers = self.paid_workers_entry.get().split(', ')
        self.volunteers = self.volunteers_entry.get().split(', ')
        grgr = self.radio0.get()
        w_tickets = self.radio1.get()
        v_tickets = self.radio2.get()
        freeplay = self.radio4.get()

        if freeplay:
            start = 10
            end = 18
        else:
            start = 10
            end = 17

        times = pd.to_datetime([datetime.time(h, m).strftime('%H:%M') for h in range(start, end) for m in (0, 30)], format='%H:%M').strftime('%I:%M')
        if self.volunteers[0]: # if no volunteers entered, entrybox still returns empty string. So list = ['']
            self.df = pd.DataFrame(columns=self.paid_workers + self.volunteers, index=times)
        else:
            self.df = pd.DataFrame(columns=self.paid_workers, index=times)

        self.sheet_frame = sheetFrame(controller = self)
        self.sheet_frame.pack_propagate(False)
        if freeplay: # more height to accomidate 5:00 and 5:30 shift.
            self.sheet_frame.place(height = 415, width = 975, relx=.992, rely=.0125, anchor='ne')
        else:
            self.sheet_frame.place(height = 365, width = 975, relx=.992, rely=.0125, anchor='ne')

        self.inputs = inputFrame(controller = self)
        self.update_sheet()

    def open_excel(self): # not implemented yet.
        # try:
        #     open_file(RES_FILE_NAME)
        # except FileNotFoundError:
        #     messagebox.showwarning('Excel Error', 'No schedule created yet, cannot open in Excel.')
        # except PermissionError:
        #     messagebox.showwarning('Excel Error', 'Please close the current excel sheet before opening the new one.')
        # except Exception as e:
        #     messagebox.showwarning('Excel Error', e)
        return

    def close(self):
        #destroy_treeview(self)
        self.save_notes()
        self.destroy()
        #open_file(RES_FILE_NAME)
    

class sheetFrame(tk.Frame):
    def __init__(self, controller: ScheduleApp):
        super().__init__(controller)
        self.controller = controller
        self.sheet = None

    def create_sheet(self, output_df):
        self.sheet = Sheet(self,
                        data=output_df.values.tolist(),
                        headers=output_df.columns.tolist(),
                        row_index=output_df.index.tolist(),
                        auto_resize_columns=True,
                        auto_resize_rows=True,
                        empty_horizontal=True,
                        empty_vertical=True
                        )
        self.sheet.enable_bindings()
        self.sheet.disable_bindings("column_width_resize", "row_height_resize", "move_columns", "move_rows", "column_height_resize", "row_width_resize")
        self.sheet.pack(fill="both", expand=True)
        self.sheet.readonly_columns(columns=[i for i, _ in enumerate(output_df.columns)], readonly=True)


    def update_sheet(self):
        """Update the display sheet with the information in the dataframe."""
        try:
            if not self.sheet:
                # This method of if else allows the sheet to not flicker upon every update.
                # The fillna('') is only to display the gray cells instead of nan.
                self.create_sheet(self.controller.df.fillna(""))
            else:
                self.sheet.set_sheet_data(data=self.controller.df.fillna("").values.tolist())

            self.color_format()

        except tk.TclError:
            pass

    def destroy_sheet(self):
        """Destroys sheets."""
        for child in self.winfo_children():
            if isinstance(child, Sheet):
                child.destroy()

    def color_format(self):
        """Apply color formatting to sheet."""
        for row_num, row in enumerate(self.sheet):
            for col_num, shift in enumerate(row):
                if pd.isna(shift):
                    self.sheet.highlight_cells(row=row_num, column=col_num, bg=None)
                elif shift in CELL_COLORS:
                    self.sheet.highlight_cells(row=row_num, column=col_num, bg=CELL_COLORS[shift])
                else:
                    self.sheet.highlight_cells(row=row_num, column=col_num, bg=None)

class inputFrame(tk.Frame):
    def __init__(self, controller: ScheduleApp):
        super().__init__(controller)
        self.controller = controller
        self.configure(background='lightblue')
        self.place(relx= .5 + 240/1500, y = 438, anchor='ne')
        self.create_widgets()

    def create_widgets(self):
        self.nonstandardFrame = NonStandardShiftFrame(self,self.controller)
        self.standardFrame = StandardShiftFrame(self,self.controller)
        self.nonstandardFrame.grid(row=0, column=1, columnspan=2, sticky="ne", pady=0, padx=4)
        self.standardFrame.grid(row=0,column=0,columnspan=1, sticky="w", pady=0)

        self.undo_button = tk.Button(self, text="Undo", command=self.controller.undo, width=13, height=2, foreground="#FFFFFF") 
        self.undo_button.config(background=secondary_button_color) 
        self.undo_button.grid(row=0, column=1, columnspan=1, sticky='e', pady=1, padx=1)
        self.undo_button.bind('<Enter>', lambda e: self.undo_button.configure(background=secondary_button_hover_color))
        self.undo_button.bind('<Leave>', lambda e: self.undo_button.configure(background=secondary_button_color))

        self.redo_button = tk.Button(self, text="Redo", command=self.controller.redo, width=13, height=2, foreground="#FFFFFF")
        self.redo_button.config(background=secondary_button_color)
        self.redo_button.grid(row=0, column=2, columnspan=1, sticky="w", pady=1, padx=1)
        self.redo_button.bind('<Enter>', lambda e: self.redo_button.configure(background=secondary_button_hover_color))
        self.redo_button.bind('<Leave>', lambda e: self.redo_button.configure(background=secondary_button_color))



class NonStandardShiftFrame(tk.Frame):
    def __init__(self, parent: inputFrame, controller: ScheduleApp):
        super().__init__(parent, relief=tk.GROOVE, borderwidth=2)
        self.parent = parent
        self.controller = controller
        self.create_widgets()

    def create_widgets(self):
        label = tk.Label(self, text="Add nonstandard shift")
        label.grid(row=0, column=0, columnspan=3, sticky="w", pady=5)

        self.entry = tk.Entry(self, width=20)
        self.entry.grid(row=1, column=0, padx=5, pady=5)

        add_button = tk.Button(self, text="Add Shift", command=self.add_shift_action, background=primary_button_color)
        add_button.grid(row=1, column=1, padx=5, pady=5)
        add_button.bind('<Enter>', lambda e: add_button.configure(background=primary_button_hover_color))
        add_button.bind('<Leave>', lambda e: add_button.configure(background=primary_button_color))


    def add_shift_action(self):
        """Activates when 'Add Shift' button is pressed."""
        selection = self.controller.get_sheet_selection()
        if not selection:
            return
        shift = self.entry.get()
        if shift == 'DELETE': 
            shift = np.nan
        selection['shift'] = shift

        self.controller.add_nonstandard_shift(selection)


class StandardShiftFrame(tk.Frame):
    def __init__(self, parent: inputFrame, controller: ScheduleApp):
        super().__init__(parent, relief=tk.GROOVE, borderwidth=2)
        self.parent = parent
        self.controller = controller
        self.create_widgets()

    def create_widgets(self):
        label = tk.Label(self, text="Add standard shift")
        label.grid(row=0, column=0, columnspan=3, sticky="w", pady=5)

        self.listbox = tk.Listbox(self, selectmode='single',width=20)
        self.listbox.grid(row=1, column=0, padx=5, pady=5)
        self.listbox.insert(tk.END, "Trike")
        self.listbox.insert(tk.END, "Gallery")
        self.listbox.insert(tk.END, "Front")
        self.listbox.insert(tk.END, "Back")
        self.listbox.insert(tk.END, "STST")
        self.listbox.insert(tk.END, "ENCA")
        self.listbox.insert(tk.END, "DIAR")
        self.listbox.insert(tk.END, "MOT")
        self.listbox.insert(tk.END, "Float")

        add_button = tk.Button(self, text="Add Shift", command=self.add_standard_action, background=primary_button_color, height=10)
        add_button.grid(row=1, column=1, padx=5, pady=5)
        add_button.bind('<Enter>', lambda e: add_button.configure(background=primary_button_hover_color))
        add_button.bind('<Leave>', lambda e: add_button.configure(background=primary_button_color))

    def add_standard_action(self):
        """Activates when 'Add shift' button is pressed."""
        shift = self.listbox.get(tk.ACTIVE)
        self.controller.save_state()
        self.controller.add_standard_shift(shift)


def show_ui() -> None:
    """Creates a GUI."""
    app = ScheduleApp()
    app.protocol('WM_DELETE_WINDOW', app.close)
    app.mainloop()

if __name__ == '__main__':
    show_ui()
    '''
    Final commit added comments so you can pull what you want out.

    Elements unfinished:
    - logic for lunches, security, tickets.
    - excel print
    - Nonstandard_history saving upon pressing "Create Schedule" a second time.
    '''
    #open_file(RES_FILE_NAME)