import sys
import pandas as pd
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout,
    QTableWidget, QTableWidgetItem, QPushButton, QMessageBox,
    QFileDialog, QHBoxLayout, QMenuBar, QLabel, QAction,
    QTextEdit, QSplitter
)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QColor, QFont


class GoogleSheetsViewer(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Visualizador de Planilha Google")
        self.setGeometry(100, 100, 1100, 700)

        self.sheet_csv_url = (
            "https://docs.google.com/spreadsheets/d/1XuGWm_gDG5edw9YkznTQGABTBah1Ptz9lfstoFdGVbA/export?format=csv&gid=49303292"
        )

        self.dataframe = pd.DataFrame()
        self.selected_region = None

        self.init_ui()
        self.load_data()

    def init_ui(self):
        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)

        main_layout = QHBoxLayout(self.central_widget)

        left_widget = QWidget()
        left_layout = QVBoxLayout(left_widget)

        menubar = QMenuBar(self)
        self.setMenuBar(menubar)
        self.region_menu = menubar.addMenu("Filtrar por Região")
        self.author_menu = menubar.addMenu("Filtrar por Autor")

        self.region_label = QLabel("Visualizando: Geral (Todas)")
        self.region_label.setMargin(5)
        left_layout.addWidget(self.region_label)

        self.metrics_label = QLabel("")
        self.metrics_label.setMargin(5)
        self.metrics_label.setStyleSheet("font-style: italic; color: gray;")
        left_layout.addWidget(self.metrics_label)

        self.table = QTableWidget()
        self.table.verticalHeader().setVisible(False)
        self.table.setAlternatingRowColors(True)
        self.table.setSortingEnabled(True)
        self.table.setSelectionBehavior(QTableWidget.SelectItems)
        self.table.horizontalHeader().setStretchLastSection(True)
        left_layout.addWidget(self.table)

        self.table.cellClicked.connect(self.show_row_details_in_console)

        button_layout = QHBoxLayout()
        self.reload_button = QPushButton("🔄 Recarregar Dados")
        self.reload_button.clicked.connect(self.load_data)
        button_layout.addWidget(self.reload_button)

        self.export_button = QPushButton("📀 Exportar CSV")
        self.export_button.clicked.connect(self.export_to_csv)
        button_layout.addWidget(self.export_button)

        left_layout.addLayout(button_layout)

        right_widget = QWidget()
        right_layout = QVBoxLayout(right_widget)

        console_label = QLabel("Preview do Terminal")
        self.clear_console_button = QPushButton("🧹 ")
        
        console_label.setAlignment(Qt.AlignCenter)
        console_label.setStyleSheet("font-weight: bold;")
        right_layout.addWidget(console_label)

        self.clear_console_button.clicked.connect(lambda: (self.console.clear(), self.append_console("Terminal limpo.")))
        right_layout.addWidget(self.clear_console_button)

        self.console = QTextEdit()
        self.console.setReadOnly(True)
        self.console.setStyleSheet("background-color: black; color: white; font-family: monospace;")
        right_layout.addWidget(self.console)

        splitter = QSplitter(Qt.Horizontal)
        splitter.addWidget(left_widget)
        splitter.addWidget(right_widget)
        splitter.setStretchFactor(0, 7)
        splitter.setStretchFactor(1, 3)

        main_layout.addWidget(splitter)

    def load_data(self):
        try:
            df = pd.read_csv(self.sheet_csv_url)
            df = self.expand_ref_column(df)

            self.dataframe = df
            self.update_region_menu()
            self.update_author_menu()
            self.populate_table()
            self.update_metrics()
            self.append_console("Dados carregados com sucesso.")
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Erro ao carregar dados:\n{e}")
            self.append_console(f"Erro ao carregar dados: {e}")

    def append_console(self, message):
        self.console.append(message)

    def expand_ref_column(self, df):
        if "Ref" not in df.columns:
            return df

        new_rows = []

        for _, row in df.iterrows():
            ref_cell = row["Ref"]
            if pd.isna(ref_cell):
                new_rows.append(row)
                continue

            parts = str(ref_cell).split("\n\n")

            if len(parts) == 1:
                new_rows.append(row)
            else:
                first_part = parts[0].strip()
                new_row = row.copy()
                new_row["Ref"] = first_part
                new_rows.append(new_row)

                for part in parts[1:]:
                    new_row = pd.Series(index=df.columns)
                    for col in df.columns:
                        new_row[col] = part.strip() if col == "Ref" else ""
                    new_rows.append(new_row)

        return pd.DataFrame(new_rows).reset_index(drop=True)

    def update_region_menu(self):
        self.region_menu.clear()
        self.selected_region = None
        action_todas = QAction("Todas", self)
        action_todas.triggered.connect(lambda: self.filtrar_por_regiao(None))
        self.region_menu.addAction(action_todas)

        if "Region" not in self.dataframe.columns:
            return

        regioes = sorted(self.dataframe["Region"].dropna().unique())

        for regiao in regioes:
            action = QAction(regiao, self)
            action.triggered.connect(lambda checked, r=regiao: self.filtrar_por_regiao(r))
            self.region_menu.addAction(action)
    def update_author_menu(self):
        self.author_menu.clear()
        # Item para desmarcar todos
        action_todos = QAction("Todos", self)
        action_todos.triggered.connect(self.clear_autor_filters)
        self.author_menu.addAction(action_todos)
        self.author_menu.addSeparator()

        if "Autores" not in self.dataframe.columns:
            return

        autores_series = self.dataframe["Autores"].dropna().apply(lambda x: [a.strip() for a in str(x).split(",")])
        todos_autores = [autor for sublist in autores_series for autor in sublist]
        contagem_autores = pd.Series(todos_autores).value_counts()

        self.autor_actions = []  # armazenar ações para controle

        for autor, count in contagem_autores.items():
            action = QAction(f"{autor} ({count})", self)
            action.setCheckable(True)
            action.toggled.connect(self.autor_selection_changed)
            action.setData(autor)  # guardar o nome do autor na ação
            self.author_menu.addAction(action)
            self.autor_actions.append(action)

    def clear_autor_filters(self):
        # desmarca todos os autores e mostra tudo
        for action in getattr(self, "autor_actions", []):
            action.setChecked(False)
        self.filtrar_por_autores([])

    def autor_selection_changed(self, checked):
        # coleta autores marcados e filtra
        autores_selecionados = [
            action.data() for action in getattr(self, "autor_actions", [])
            if action.isChecked()
        ]
        self.filtrar_por_autores(autores_selecionados)

    def filtrar_por_autores(self, autores):
        if not autores:
            # Nenhum autor selecionado, mostrar tudo
            self.region_label.setText("Visualizando: Geral (Todas)")
            self.populate_table()
            self.append_console("Filtro de autor removido, mostrando todos.")
            return

        df = self.dataframe.copy()

        def contem_autor(cell):
            if pd.isna(cell):
                return False
            lista_autores = [a.strip() for a in str(cell).split(",")]
            return any(autor in lista_autores for autor in autores)

        df_filtrado = df[df["Autores"].apply(contem_autor)]

        self.region_label.setText(f"Visualizando: Autor(es) - {', '.join(autores)}")
        self.populate_table_custom(df_filtrado)
        self.append_console(f"Filtro aplicado: Autor(es) - {', '.join(autores)}")



    def filtrar_por_regiao(self, regiao):
        self.selected_region = regiao
        self.region_label.setText(f"Visualizando: {regiao if regiao else 'Geral (Todas)'}")
        self.populate_table()
        self.append_console(f"Filtro aplicado: {regiao if regiao else 'Todas as regiões'}")

    def filtrar_por_autor(self, autor):
        if autor is None:
            # Limpar filtro de autor e mostrar todos
            self.region_label.setText("Visualizando: Geral (Todas)")
            self.selected_region = None
            self.populate_table()
            self.append_console("Filtro de autor removido, mostrando todos.")
            return

        df = self.dataframe.copy()
        df = df[df["Autores"].fillna("").apply(lambda x: autor in [a.strip() for a in str(x).split(",")])]
        self.region_label.setText(f"Visualizando: Autor - {autor}")
        self.selected_region = None
        self.populate_table_custom(df)
        self.append_console(f"Filtro aplicado: Autor - {autor}")

    def populate_table_custom(self, df):
        self.table.clear()
        self.table.setRowCount(len(df))
        self.table.setColumnCount(len(df.columns))
        self.table.setHorizontalHeaderLabels(df.columns)

        for row in range(len(df)):
            for col, column_name in enumerate(df.columns):
                value = str(df.iat[row, col])
                item = QTableWidgetItem(value)
                item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                self.table.setItem(row, col, item)

        self.update_spans()
        self.color_region_blocks()
        self.update_metrics()

    def clear_spans(self):
        for r in range(self.table.rowCount()):
            for c in range(self.table.columnCount()):
                if self.table.rowSpan(r, c) > 1 or self.table.columnSpan(r, c) > 1:
                    self.table.setSpan(r, c, 1, 1)

    def populate_table(self):
        df = self.dataframe
        if self.selected_region:
            df = df[df["Region"] == self.selected_region]

        self.table.clear()
        self.table.setRowCount(len(df))
        self.table.setColumnCount(len(df.columns))
        self.table.setHorizontalHeaderLabels(df.columns)

        font_bold = QFont()
        font_bold.setBold(True)

        region_colors = {
            "Africa": QColor(240, 230, 140),
            "Asia": QColor(173, 216, 230),
            "Australia and New Zeland": QColor(255, 235, 205),
            "Canada and EUA": QColor(200, 200, 255),
            "Europe": QColor(221, 160, 221),
            "Latin America": QColor(152, 251, 152),
            "Eastern Mediterranean": QColor(255, 182, 193),
        }

        for row in range(len(df)):
            region = df.iloc[row].get("Region", "").strip() if "Region" in df.columns else ""
            bg_color = region_colors.get(region, QColor(255, 255, 255))

            for col, column_name in enumerate(df.columns):
                value = str(df.iat[row, col])
                item = QTableWidgetItem(value)
                item.setFlags(item.flags() & ~Qt.ItemIsEditable)

                if value.strip() == "":
                    item.setBackground(QColor(220, 220, 220))
                    item.setForeground(QColor(0, 0, 0))
                    item.setTextAlignment(Qt.AlignCenter)
                else:
                    item.setBackground(bg_color)
                    item.setForeground(QColor(0, 0, 0))

                if column_name.lower() == "num":
                    item.setFont(font_bold)
                    item.setTextAlignment(Qt.AlignCenter)

                self.table.setItem(row, col, item)

        self.update_spans()
        self.color_region_blocks()
        self.update_metrics()

    def color_region_blocks(self):
        try:
            region_col_index = self.dataframe.columns.get_loc("Region")
        except KeyError:
            return

        block_color_1 = QColor(255, 255, 255)
        block_color_2 = QColor(245, 245, 245)

        last_value = None
        block_num = 0

        for row in range(self.table.rowCount()):
            item = self.table.item(row, region_col_index)
            if item is None:
                continue

            current_value = item.text().strip()
            if current_value != last_value:
                block_num += 1
                last_value = current_value

            block_color = block_color_1 if (block_num % 2) == 1 else block_color_2
            for col in range(self.table.columnCount()):
                cell = self.table.item(row, col)
                if cell:
                    bg = cell.background().color()
                    blended = QColor(
                        (bg.red() + block_color.red()) // 2,
                        (bg.green() + block_color.green()) // 2,
                        (bg.blue() + block_color.blue()) // 2,
                    )
                    cell.setBackground(blended)

    def update_spans(self):
        self.clear_spans()
        header_labels = [self.table.horizontalHeaderItem(i).text() for i in range(self.table.columnCount())]
        columns_conditional_merge = [
            "Num", "Region", "country", "Autores", "Titulo", "Afiliation", "Resumo",
            "Num de Ref", "IA abstract 100 palavras", "IA keywords"
        ]

        for col_name in columns_conditional_merge:
            if col_name not in header_labels or col_name == "Ref":
                continue

            col = header_labels.index(col_name)
            row = 0

            while row < self.table.rowCount():
                item = self.table.item(row, col)
                if item is None or item.text().strip() == "":
                    row += 1
                    continue

                start_row = row
                span_count = 1
                next_row = row + 1

                while next_row < self.table.rowCount():
                    next_item = self.table.item(next_row, col)
                    if next_item and next_item.text().strip() == "":
                        span_count += 1
                        next_row += 1
                    else:
                        break

                if span_count > 1:
                    self.table.setSpan(start_row, col, span_count, 1)

                row = next_row

    def update_metrics(self):
        df = self.dataframe
        if self.selected_region:
            df = df[df["Region"] == self.selected_region]

        if "Num" in df.columns:
            total_com_num = df["Num"].dropna().astype(str).str.strip().replace("", pd.NA).dropna().shape[0]
            self.metrics_label.setText(f"Abstracts: {total_com_num}")
        else:
            self.metrics_label.setText("Coluna 'Num' não encontrada.")

    def export_to_csv(self):
        headers = [self.table.horizontalHeaderItem(i).text() for i in range(self.table.columnCount())]
        if not headers:
            self.append_console("Nada para exportar.")
            return

        rows = []
        for row in range(self.table.rowCount()):
            row_data = []
            for col in range(self.table.columnCount()):
                item = self.table.item(row, col)
                row_data.append(item.text() if item else "")
            rows.append(row_data)

        df_export = pd.DataFrame(rows, columns=headers)

        file_path, _ = QFileDialog.getSaveFileName(self, "Salvar CSV", "", "CSV Files (*.csv)")
        if file_path:
            try:
                df_export.to_csv(file_path, index=False)
                self.append_console(f"Arquivo salvo em: {file_path}")
            except Exception as e:
                QMessageBox.critical(self, "Erro", f"Erro ao salvar arquivo:\n{e}")
                self.append_console(f"Erro ao salvar arquivo: {e}")

    def show_row_details_in_console(self, row, column):
        # Imprime dados da linha clicada no console para facilitar a visualização
        if row < 0 or row >= self.table.rowCount():
            return

        dados = []
        for col in range(self.table.columnCount()):
            header = self.table.horizontalHeaderItem(col).text()
            item = self.table.item(row, col)
            valor = item.text() if item else ""
            dados.append(f"{header}: {valor}")

        self.append_console("Detalhes da linha selecionada:\n" + "\n".join(dados))


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = GoogleSheetsViewer()
    window.show()
    sys.exit(app.exec_())
