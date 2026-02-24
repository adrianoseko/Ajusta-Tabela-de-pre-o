import os
import sys
import logging
from typing import Any, Optional, List

import pandas as pd
import pyodbc
from tkinter import Tk, StringVar, END, NORMAL, DISABLED
from tkinter import messagebox
from tkinter import Button, Entry, Label, OptionMenu
from tkinter import ttk


logging.basicConfig(level=logging.INFO)


class PriceTableApp:
    """GUI application to view and edit price table items.

    Preserves original application behavior while improving structure and clarity.
    """

    DB_DRIVER = "{SQL Server}"
    DB_SERVER = " Seu Servidor"
    DB_UID = "login"
    DB_PWD = "senha"

    def __init__(self, argv: List[str]):
        # Parse arguments (kept consistent with original positions)
        try:
            self.bc = argv[1]
            self.us = argv[2]
            self.cl = argv[3]
        except IndexError as exc:
            raise SystemExit("Expected three command-line arguments: banco, usuario, coligada") from exc

        # Extract tokens from incoming args (same split logic as original)
        self.coligada = self.cl.split('/c:')
        self.banco = self.bc.split('/d:')
        self.user = self.us.split('/u:')

        # DB driver and connection (kept as in original)
        self.drive = self.DB_DRIVER
        self.cnxn = self._connect_db()
        self.cursor = self.cnxn.cursor()

        # Data containers
        self.tables_df: pd.DataFrame = pd.DataFrame()
        self.items_df: pd.DataFrame = pd.DataFrame()
        self.selected_table_index: Optional[int] = None
        self.selected_table_id: Optional[Any] = None

        # Build UI
        self.root = Tk()
        self.root.geometry("800x600")
        self.root.title("CGA.NET - Tabela Preços")

        # Tk variables
        self.selected_table_var = StringVar(self.root)
        self.filter_var = StringVar(self.root)

        # Create widgets
        self._create_widgets()

        # Load tables for selection
        self._load_table_list()

        # Trace selection change (mimics original var.trace("w", callback))
        self.selected_table_var.trace("w", lambda *args: self.on_table_selected())

        # Start mainloop
        self.root.mainloop()

    def _connect_db(self) -> pyodbc.Connection:
        try:
            conn_str = (
                f"DRIVER={self.drive};SERVER={self.DB_SERVER};DATABASE={self.banco[1]};"
                f"UID={self.DB_UID};PWD={self.DB_PWD};"
            )
            return pyodbc.connect(conn_str)
        except Exception as exc:
            logging.exception("Failed to connect to database")
            messagebox.showerror("Erro DB", "Não foi possível conectar ao banco de dados.\nEntre em contato com o suporte!")
            raise

    def _create_widgets(self) -> None:
        # Option menu for tables
        self.table_menu = OptionMenu(self.root, self.selected_table_var, "")
        self.table_menu.config(width=30, font=("Helvetica", 12))
        self.table_menu.place(x=20, y=10)

        # Treeview for items
        self.tree = ttk.Treeview(
            self.root,
            selectmode='extended',
            column=(
                'Column1', 'Column2', 'Column3', 'Column4', 'Column5', 'Column6', 'Column7'
            ),
            show='headings',
            height=20,
        )

        self.vscrollbar = ttk.Scrollbar(self.root, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=self.vscrollbar.set)

        # Configure columns (keep same widths and headings)
        self.tree.column('Column1', width=80, minwidth=50, stretch=False)
        self.tree.heading("#1", text="ID Produto")

        self.tree.column('Column2', width=100, minwidth=50, stretch=False)
        self.tree.heading("#2", text="Código Produto")

        self.tree.column('Column3', width=300, minwidth=100, stretch=False)
        self.tree.heading("#3", text="Nome Produto")

        self.tree.column('Column4', width=60, minwidth=50, stretch=False)
        self.tree.heading("#4", text="Preço")

        self.tree.column('Column5', width=60, minwidth=50, stretch=False)
        self.tree.heading("#5", text="Custo")

        self.tree.column('Column6', width=60, minwidth=50, stretch=False)
        self.tree.heading("#6", text="Margem")

        self.tree.column('Column7', width=100, minwidth=50, stretch=False)
        self.tree.heading("#7", text="Adcional Financeiro")

        # Bind double click on treeview to edit
        self.tree.bind("<Double-1>", lambda e: self.select_item())

        # Buttons
        self.save_button = Button(self.root, text="Gravar", command=self.record, height=1, width=10)
        self.cancel_button = Button(self.root, text="Cancelar", command=self.cancel_edit, height=1, width=10)
        self.export_button = Button(self.root, text="Exportar", command=self.export, height=1, width=10)
        self.cancel2_button = Button(self.root, text="Cancelar", command=self.cancel_to_menu, height=1, width=10)

        # Keep them hidden initially (pack/place forget as original)
        self.save_button.pack_forget()
        self.cancel_button.pack_forget()
        self.export_button.pack_forget()
        self.cancel2_button.pack_forget()

        # Filter widgets
        self.filter_label = Label(self.root, text="Pesquisar:")
        self.filter_entry = Entry(self.root, textvariable=self.filter_var, width=50)
        # Do not place immediately (original places on selection)

        # Entry fields for edit (labels and entries)
        self.code_label = Label(self.root, text="Código do Produto")
        self.code_entry = Entry(self.root, bd=5)

        self.name_label = Label(self.root, text="Nome do Produto")
        self.name_entry = Entry(self.root, bd=5, width=70)

        self.price_label = Label(self.root, text="Preço")
        self.price_entry = Entry(self.root, bd=5)

        self.cost_label = Label(self.root, text="Custo")
        self.cost_entry = Entry(self.root, bd=5)

        self.margin_label = Label(self.root, text="Margem")
        self.margin_entry = Entry(self.root, bd=5)

        self.addf_label = Label(self.root, text="Adicional Financeiro")
        self.addf_entry = Entry(self.root, bd=5)

        # Info labels
        Label(self.root, text=f"Banco: {self.banco[1]}").place(x=50, y=580)
        Label(self.root, text=f"Coligada: {self.coligada[1]}").place(x=200, y=580)
        Label(self.root, text=f"Usuario: {self.user[1]}").place(x=280, y=580)
        Label(self.root, text=f"Versão: 1.0").place(x=500, y=580)

    def _load_table_list(self) -> None:
        sql = (
            "SELECT ZTC.IDTABPRECO, NOME from ZA_TTABPRECO ZTC "
            "LEFT JOIN TTABPRECO TTP (NOLOCK) ON TTP.CODCOLIGADA = ZTC.CODCOLIGADA AND TTP.IDTABPRECO = ZTC.IDTABPRECO "
            "where USADEFAULTABELA = 'N' and "
            "TTP.IDTABPRECO > 3 AND TTP.ATIVA = 1 AND "
            "CONVERT(VARCHAR(10) , GETDATE() , 126) >= CONVERT(VARCHAR(10) , TTP.DATAVIGENCIAINI , 126) AND"
            "CONVERT(VARCHAR(10) , GETDATE() , 126) <= CONVERT(VARCHAR(10) , TTP.DATAVIGENCIAFIM , 126)"
        )

        try:
            self.tables_df = pd.read_sql(sql, self.cnxn)
            if self.tables_df.empty:
                messagebox.showinfo("Informação", "Nenhuma tabela de preço encontrada.")
                return

            names = self.tables_df['NOME'].values.tolist()
            # Update OptionMenu with names
            menu = self.table_menu['menu']
            menu.delete(0, 'end')
            for name in names:
                menu.add_command(label=name, command=lambda v=name: self.selected_table_var.set(v))

            # Default selection like original: set to first
            self.selected_table_var.set(names[0])
        except Exception:
            logging.exception("Failed to load table list")
            messagebox.showerror("ERRO!", '"Não foi possível obter lista de tabelas!\nEntre em contato com o suporte!"')

    def on_table_selected(self) -> None:
        """Callback when a table is selected from the OptionMenu."""
        selected_name = str(self.selected_table_var.get())
        # Clear tree
        self.tree.delete(*self.tree.get_children())

        # Find index of selected table name (replicates original loop behavior)
        try:
            names = self.tables_df['NOME'].values.tolist()
            ids = self.tables_df['IDTABPRECO'].values.tolist()
            index = names.index(selected_name)
            self.selected_table_index = index
            self.selected_table_id = ids[index]
        except Exception:
            logging.exception("Selected table name not found in table list")
            return

        # Query items for selected table
        sql2 = (
            "SELECT  IDTABPRECO, ZTC.IDPRD , CODIGOPRD, NOMEFANTASIA, PRECO, CUSTO, MARGEM, ADIC_FINANC  from ZA_TTABPRECOITM ZTC "
            "LEFT JOIN TPRD (NOLOCK) ON TPRD.CODCOLIGADA = ZTC.CODCOLIGADA AND TPRD.IDPRD = ZTC.IDPRD "
            f"where IDTABPRECO = {self.selected_table_id} ORDER BY NOMEFANTASIA"
        )

        try:
            self.items_df = pd.read_sql(sql2, self.cnxn)
        except Exception:
            logging.exception("Failed to read items for table")
            messagebox.showerror("ERRO!", '"Não foi possível carregar itens da tabela!\nEntre em contato com o suporte!"')
            return

        # Populate treeview with rows
        ids = self.items_df['IDPRD'].values.tolist()
        codes = self.items_df['CODIGOPRD'].values.tolist()
        names = self.items_df['NOMEFANTASIA'].values.tolist()
        prices = self.items_df['PRECO'].values.tolist()
        costs = self.items_df['CUSTO'].values.tolist()
        margins = self.items_df['MARGEM'].values.tolist()
        addfs = self.items_df['ADIC_FINANC'].values.tolist()

        for i in range(len(codes)):
            self.tree.insert('', END, values=(ids[i], codes[i], names[i], prices[i], costs[i], margins[i], addfs[i]), tag=str(i))

        # Place UI elements (mimicking original placement)
        self.filter_var.trace("w", lambda *args: self.apply_filter())
        self.tree.place(x=20, y=50)
        self.export_button.place(x=700, y=500)
        self.cancel2_button.place(x=700, y=550)
        self.vscrollbar.place(x=783, y=50, height=425)
        self.filter_entry.place(x=450, y=15)
        self.filter_label.place(x=390, y=15)

    def select_item(self) -> None:
        """Prepare editing UI for the selected row."""
        try:
            # Hide some widgets like original
            self.cancel2_button.place_forget()
            self.export_button.place_forget()
            self.table_menu.place_forget()
            self.filter_entry.place_forget()
            self.filter_label.place_forget()

            # Disable tree scrolling while editing (original used configure with None)
            self.vscrollbar.place(x=783, y=50, height=425)
            self.tree.configure(yscrollcommand=None)

            # Show edit fields
            self._show_edit_widgets()

            # Get selected item values
            selected_id = self.tree.selection()[0]
            item_values = self.tree.item(selected_id, "values")

            # Fill entries like original and disable code/name editing
            self.code_entry.delete(0, END)
            self.code_entry.insert(0, item_values[1])
            self.code_entry['state'] = DISABLED

            self.name_entry.delete(0, END)
            self.name_entry.insert(0, item_values[2])
            self.name_entry['state'] = DISABLED

            self.price_entry.delete(0, END)
            self.price_entry.insert(0, item_values[3])

            self.cost_entry.delete(0, END)
            self.cost_entry.insert(0, item_values[4])

            self.margin_entry.delete(0, END)
            self.margin_entry.insert(0, item_values[5])

            self.addf_entry.delete(0, END)
            self.addf_entry.insert(0, item_values[6])

            # Disable tree interactions while editing
            self.tree.state(("disabled",))
            self.tree.bind('<Button-1>', lambda e: 'break')

        except Exception:
            logging.exception("Error preparing item for edit")
            messagebox.showerror("Erro", '"Não foi possível selecionar item para edição.\nEntre em contato com o suporte!"')

    def _show_edit_widgets(self) -> None:
        # Place edit widgets in the same coordinates as original
        self.code_entry['state'] = NORMAL
        self.code_label.place(x=50, y=480)
        self.code_entry.place(x=50, y=500)

        self.name_entry['state'] = NORMAL
        self.name_label.place(x=200, y=480)
        self.name_entry.place(x=200, y=500)

        self.price_label.place(x=50, y=530)
        self.price_entry.place(x=50, y=550)

        self.cost_label.place(x=200, y=530)
        self.cost_entry.place(x=200, y=550)

        self.margin_label.place(x=350, y=530)
        self.margin_entry.place(x=350, y=550)

        self.addf_label.place(x=500, y=530)
        self.addf_entry.place(x=500, y=550)

        self.save_button.place(x=700, y=530)
        self.cancel_button.place(x=700, y=560)

    def _hide_edit_widgets(self) -> None:
        # Hide edit widgets (used in multiple places)
        self.code_label.place_forget()
        self.code_entry.place_forget()
        self.name_label.place_forget()
        self.name_entry.place_forget()
        self.price_label.place_forget()
        self.price_entry.place_forget()
        self.cost_label.place_forget()
        self.cost_entry.place_forget()
        self.margin_label.place_forget()
        self.margin_entry.place_forget()
        self.addf_label.place_forget()
        self.addf_entry.place_forget()
        self.save_button.place_forget()
        self.cancel_button.place_forget()

    def apply_filter(self) -> None:
        """Filter treeview rows by product name (case sensitive by original logic using upper)."""
        try:
            item_text = self.filter_var.get().upper()
            # Iterate over current children and reinsert matching ones at top if matches
            children = list(self.tree.get_children())
            for child in children:
                values = self.tree.item(child)['values']
                # values[2] is product name (NOMEFANTASIA)
                if item_text in str(values[2]):
                    # Move it to the top (original removed and re-inserted at position 0)
                    found = values
                    self.tree.delete(child)
                    self.tree.insert('', 0, values=found)
        except Exception:
            logging.exception("Filter failed")
            messagebox.showerror("ERRO!", '"Não foi possível obter filtro!\nEntre em contato com o suporte!"')

    def record(self) -> None:
        """Persist edited values to the database and refresh the view."""
        try:
            selected_id = self.tree.selection()[0]
            row_values = self.tree.item(selected_id, "values")
            idprd = row_values[0]

            preco = self.price_entry.get()
            custo = self.cost_entry.get()
            margem = self.margin_entry.get()
            adc = self.addf_entry.get()

            # Clear edit fields
            self.price_entry.delete(0, END)
            self.cost_entry.delete(0, END)
            self.margin_entry.delete(0, END)
            self.addf_entry.delete(0, END)

            # Update SQL (kept same structure as original)
            update = (
                "UPDATE ZA_TTABPRECOITM SET PRECO = {preco}, CUSTO = {custo}, MARGEM = {margem}, ADIC_FINANC = {adc}"
                " WHERE CODCOLIGADA = 5 AND IDPRD = {idprd} and IDTABPRECO = {idtab}"
            ).format(preco=preco, custo=custo, margem=margem, adc=adc, idprd=idprd, idtab=self.selected_table_id)

            self._execute_update(update)

            # Hide edit fields and restore UI
            self._hide_edit_widgets()
            self.vscrollbar.place_forget()

            messagebox.showinfo("Sucesso!", "Alterações salvas com sucesso!")

            # Refresh table contents (simulates callback)
            self.on_table_selected()
        except Exception:
            logging.exception("Failed to save changes")
            messagebox.showerror("Erro Sistema", '"Não foi possível efetuar as alterações \b\n                             Entre em contato com o suporte!"')

    def _execute_update(self, sql_statement: str) -> None:
        try:
            self.cursor.execute(sql_statement)
            # commit via connection to ensure persistence
            self.cnxn.commit()
        except Exception:
            logging.exception("Database update failed")
            raise

    def export(self) -> None:
        """Export current items DataFrame to Excel and open it."""
        try:
            if self.items_df.empty:
                messagebox.showinfo("Exportação Excel", "Nada para exportar")
                return

            filename = f"{self.tables_df['NOME'].values[self.selected_table_index]}.xlsx"
            self.items_df.to_excel(filename, index=False)
            messagebox.showinfo("Exportação Excel", "Tabela exportada com sucesso!")
            # Open hardcoded path as original did (kept intact)
            os.startfile(f"C:/Users/admin/Desktop/Adriano/Projetos_SW/TABELA DE PREÇO/{filename}")
        except Exception:
            logging.exception("Export failed")
            messagebox.showerror("Exportação Excel", '"Não foi possível exportar a tabela \b\n                             Entre em contato com o suporte!"')

    def cancel_edit(self) -> None:
        """Cancel current edit and restore UI to selectable state."""
        try:
            # Restore export/cancel buttons and filter controls like original
            self.export_button.place(x=700, y=500)
            self.cancel2_button.place(x=700, y=550)
            self.vscrollbar.place(x=783, y=50, height=425)
            self.filter_entry.place(x=450, y=15)
            self.filter_label.place(x=390, y=15)

            # Hide edit fields
            self._hide_edit_widgets()

            # Clear entries
            self.code_entry.delete(0, END)
            self.name_entry.delete(0, END)
            self.price_entry.delete(0, END)
            self.cost_entry.delete(0, END)
            self.margin_entry.delete(0, END)
            self.addf_entry.delete(0, END)

            # Re-enable tree
            self.tree.state(("!disabled",))
            self.tree.unbind('<Button-1>')
        except Exception:
            logging.exception("Cancel edit failed")
            messagebox.showerror("Erro!", "Entre em contato com o suporte!")

    def cancel_to_menu(self) -> None:
        """Cancel viewing items and return to menu state."""
        # Mimics original cancela2() behavior
        self.table_menu.place(x=20, y=10)
        self.tree.delete(*self.tree.get_children())
        self.tree.place_forget()
        self.export_button.place_forget()
        self.cancel2_button.place_forget()
        self.vscrollbar.place_forget()
        self.filter_entry.place_forget()
        self.filter_label.place_forget()


if __name__ == "__main__":
    # Instantiate application with sys.argv (keeps original behavior)
    PriceTableApp(sys.argv)
