import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import matplotlib.font_manager as fm
import tkinter as tk
from tkinter import filedialog, messagebox, ttk, colorchooser
import os
from PIL import Image, ImageTk

# --- フォント検索関数 (変更なし) ---
def find_japanese_font():
    font_paths = fm.findSystemFonts(fontpaths=None, fontext='ttf')
    japanese_fonts = set()
    for path in font_paths:
        try:
            prop = fm.FontProperties(fname=path)
            font_name = prop.get_name()
            if 'japanese' in font_name.lower() or 'gothic' in font_name.lower() or 'meiryo' in font_name.lower() or \
               'yu gothic' in font_name.lower() or 'hiragino' in font_name.lower() or 'noto sans cjk' in font_name.lower() or \
               'ms mincho' in font_name.lower() or 'ms gothic' in font_name.lower():
                japanese_fonts.add(font_name)
        except Exception:
            continue
    return sorted(list(japanese_fonts))

# --- メッコ図描画ロジック関数 (変更なし) ---
def create_mekko_chart(csv_file_path, hanrei, graph_title, xlabel, ylabel, jiku, save_dir, save_filename, font_name_setting, category_colors):
    final_font_name = None
    if font_name_setting == "Yu Gothic" or font_name_setting == "Meiryo":
        final_font_name = font_name_setting
    elif font_name_setting == "Auto":
        available_fonts = find_japanese_font()
        if "Yu Gothic" in available_fonts:
            final_font_name = "Yu Gothic"
        elif "Meiryo" in available_fonts:
            final_font_name = "Meiryo"
        elif available_fonts:
            final_font_name = available_fonts[0]
        else:
            messagebox.showwarning("警告", "日本語フォントが自動検出できませんでした。システムのデフォルトフォントを使用します。")
            final_font_name = None
    else:
        final_font_name = font_name_setting

    full_save_path = os.path.join(save_dir, save_filename)

    try:
        if final_font_name:
            plt.rcParams['font.family'] = final_font_name
        plt.rcParams['axes.unicode_minus'] = False

        columns = ["group", "type", "value"]
        df = pd.read_csv(csv_file_path, header=None, names=columns)
        if df.empty:
            raise ValueError("読み込んだCSVファイルにデータがありません。")

        df_region_total = df.groupby(columns[0])[columns[2]].sum().reset_index()
        df_region_total.rename(columns={columns[2]: 'sum_value'}, inplace=True)
        total_sales_all = df[columns[2]].sum()
        df_region_total['ratio'] = df_region_total['sum_value'] / total_sales_all
        df = pd.merge(df, df_region_total, on=columns[0], how='left')

        df['group_ratio'] = df.groupby(columns[0])[columns[2]].transform(lambda x: x / x.sum())

        df = df.sort_values(by=[columns[0], columns[1]])
        regions = df[columns[0]].unique()
        x_starts = {}
        current_x = 0
        for region in regions:
            x_starts[region] = current_x
            current_x += df_region_total[df_region_total[columns[0]] == region]['ratio'].iloc[0]

        fig, ax = plt.subplots(figsize=(10, 6))
        patches = []

        product_colors = category_colors

        for region in regions:
            region_data = df[df[columns[0]] == region]
            region_width = df_region_total[df_region_total[columns[0]] == region]['ratio'].iloc[0]
            x_start = x_starts[region]

            y_bottom = 0
            for index, row in region_data.iterrows():
                height = row['group_ratio']
                product = row[columns[1]]

                rect = mpatches.Rectangle(
                    (x_start, y_bottom),
                    width=region_width,
                    height=height,
                    facecolor=product_colors.get(product, 'gray'),
                    edgecolor='white',
                    linewidth=1.5
                )
                ax.add_patch(rect)

                ax.text(x_start + region_width / 2, y_bottom + height / 2,
                        f'{product}\n({row["group_ratio"]:.1%})',
                        ha='center', va='center', color='white', fontsize=9)
                
                y_bottom += height
            
            ax.text(x_start + region_width / 2, -0.05, region,
                    ha='center', va='top', fontsize=10, color='black')

        for product in sorted(product_colors.keys()):
            patches.append(mpatches.Patch(color=product_colors[product], label=product))

        ax.legend(handles=patches, title=hanrei, bbox_to_anchor=(1.05, 1), loc='upper left')

        ax.set_title(graph_title)
        ax.set_xlabel(xlabel)
        ax.set_ylabel(ylabel)
        ax.set_xlim(0, 1)
        ax.set_ylim(0, 1)
        ax.set_xticks([])
        ax.set_yticks([])
        ax.axis(jiku)

        plt.tight_layout()
        plt.savefig(full_save_path, dpi=300, bbox_inches='tight')
        plt.close(fig)
        return True, f"グラフを '{full_save_path}' として保存しました。"

    except FileNotFoundError as e:
        return False, f"エラー: ファイルが見つからないか、保存先フォルダが存在しません。パスを確認してください。\n詳細: {e}"
    except pd.errors.EmptyDataError:
        return False, f"エラー: ファイル '{csv_file_path}' は空です。"
    except Exception as e:
        return False, f"グラフ生成中にエラーが発生しました: {e}"

# --- GUIアプリケーションの定義 (変更あり) ---
class MekkoChartGenerator(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("メッコ図生成ツール")
        self.geometry("800x700")

        self.category_colors = {}
        self.category_color_frames = {}
        self.hint_image_refs = {}

        self.create_widgets()
        self.load_default_colors()

    def create_widgets(self):
        self.settings_frame = ttk.LabelFrame(self, text="グラフ設定", padding="10")
        self.settings_frame.pack(padx=10, pady=10, fill="x")

        labels_and_defaults = [
            ("CSVファイルパス:", ''),
            ("凡例タイトル:", ""),
            ("グラフタイトル:", ""),
            ("X軸ラベル:", ""),
            ("Y軸ラベル:", ""),
            ("保存先フォルダ:", os.getcwd()),
            ("保存ファイル名:", "mekko.png")
        ]

        self.entries = {}
        row_counter = 0
        for label_text, default_value in labels_and_defaults:
            label = ttk.Label(self.settings_frame, text=label_text)
            label.grid(row=row_counter, column=0, sticky="w", pady=2, padx=5)
            entry = ttk.Entry(self.settings_frame, width=50)
            entry.grid(row=row_counter, column=1, sticky="ew", pady=2, padx=5)
            self.entries[label_text] = entry
            entry.insert(0, default_value)

            # CSVファイルパスと保存先フォルダからはヒントボタンを削除
            if "CSVファイルパス:" == label_text:
                browse_button = ttk.Button(self.settings_frame, text="参照...", command=self.browse_csv_file)
                browse_button.grid(row=row_counter, column=2, padx=5)
            elif "保存先フォルダ:" == label_text:
                browse_dir_button = ttk.Button(self.settings_frame, text="参照...", command=self.browse_save_directory)
                browse_dir_button.grid(row=row_counter, column=2, padx=5)
            elif "凡例タイトル:" == label_text: # ここに凡例タイトルのヒントボタンを追加
                # 凡例タイトル用のヒント画像を用意（例: hint.png）
                hint_button = ttk.Button(self.settings_frame, text="ヒント",
                                         command=lambda: self.show_hint_image("凡例タイトルの例", "hint.png", (800, 500)))
                hint_button.grid(row=row_counter, column=2, padx=5) # 既存の参照ボタンの代わりに同じ列に配置、または別の列 (column=3) でも可
            
            row_counter += 1

        # 軸表示（ラジオボタン）
        axis_label = ttk.Label(self.settings_frame, text="軸表示:")
        axis_label.grid(row=row_counter, column=0, sticky="w", pady=2, padx=5)
        self.axis_display_var = tk.StringVar(value="off")
        axis_radio_frame = ttk.Frame(self.settings_frame)
        axis_radio_frame.grid(row=row_counter, column=1, sticky="w", pady=2, padx=5)
        ttk.Radiobutton(axis_radio_frame, text="On", variable=self.axis_display_var, value="on").pack(side="left", padx=5)
        ttk.Radiobutton(axis_radio_frame, text="Off", variable=self.axis_display_var, value="off").pack(side="left", padx=5)
        row_counter += 1

        # フォント選択（ラジオボタン）
        font_label = ttk.Label(self.settings_frame, text="フォント:")
        font_label.grid(row=row_counter, column=0, sticky="w", pady=2, padx=5)
        self.font_selection_var = tk.StringVar(value="Auto")
        font_radio_frame = ttk.Frame(self.settings_frame)
        font_radio_frame.grid(row=row_counter, column=1, sticky="w", pady=2, padx=5)
        ttk.Radiobutton(font_radio_frame, text="Yu Gothic", variable=self.font_selection_var, value="Yu Gothic").pack(side="left", padx=5)
        ttk.Radiobutton(font_radio_frame, text="Meiryo", variable=self.font_selection_var, value="Meiryo").pack(side="left", padx=5)
        ttk.Radiobutton(font_radio_frame, text="Auto", variable=self.font_selection_var, value="Auto").pack(side="left", padx=5)
        row_counter += 1

        # --- カテゴリ色設定フレーム (変更なし) ---
        self.category_color_frame = ttk.LabelFrame(self, text="カテゴリ色設定", padding="10")
        self.category_color_frame.pack(padx=10, pady=10, fill="both", expand=True)

        self.canvas = tk.Canvas(self.category_color_frame)
        self.scrollbar = ttk.Scrollbar(self.category_color_frame, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = ttk.Frame(self.canvas)

        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(
                scrollregion=self.canvas.bbox("all")
            )
        )

        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")

        # --- 実行ボタン (変更なし) ---
        generate_button = ttk.Button(self, text="メッコ図を生成・保存", command=self.generate_chart)
        generate_button.pack(pady=10)

        # --- メッセージ表示ラベル (変更なし) ---
        self.status_label = ttk.Label(self, text="", relief="sunken", anchor="w", padding="5")
        self.status_label.pack(padx=10, pady=5, fill="x")

    def show_hint_image(self, title, image_path, size=None):
        if not os.path.exists(image_path):
            messagebox.showerror("エラー", f"ヒント画像 '{image_path}' が見つかりません。")
            return

        try:
            hint_window = tk.Toplevel(self)
            hint_window.title(title)
            hint_window.transient(self)
            hint_window.grab_set()

            img = Image.open(image_path)
            if size:
                img = img.resize(size, Image.Resampling.LANCZOS)
            
            self.hint_image_refs[image_path] = ImageTk.PhotoImage(img) 
            
            image_label = tk.Label(hint_window, image=self.hint_image_refs[image_path])
            image_label.pack(padx=10, pady=10)

            hint_window.protocol("WM_DELETE_WINDOW", lambda: self._on_hint_window_close(hint_window, image_path))
            
            self.update_idletasks()
            parent_x = self.winfo_x()
            parent_y = self.winfo_y()
            parent_width = self.winfo_width()
            parent_height = self.winfo_height()

            hint_width = hint_window.winfo_width()
            hint_height = hint_window.winfo_height()

            x = parent_x + (parent_width // 2) - (hint_width // 2)
            y = parent_y + (parent_height // 2) - (hint_height // 2)

            hint_window.geometry(f"+{x}+{y}")

        except Exception as e:
            messagebox.showerror("エラー", f"ヒント画像の表示中にエラーが発生しました: {e}")

    def _on_hint_window_close(self, window, image_path):
        window.grab_release()
        if image_path in self.hint_image_refs:
            del self.hint_image_refs[image_path]
        window.destroy()

    def load_default_colors(self):
        csv_path = self.entries["CSVファイルパス:"].get()
        if not os.path.exists(csv_path):
            self.clear_category_color_widgets()
            return

        try:
            df = pd.read_csv(csv_path, header=None, names=["group", "type", "value"])
            unique_products = sorted(df['type'].unique())
        except Exception as e:
            messagebox.showerror("エラー", f"CSVファイルの読み込みに失敗しました: {e}")
            self.clear_category_color_widgets()
            return

        self.clear_category_color_widgets()

        import matplotlib.colors as mcolors
        default_cmap = plt.cm.get_cmap('tab10', len(unique_products))
        
        for i, product in enumerate(unique_products):
            default_color = mcolors.to_hex(default_cmap(i))

            self.category_colors[product] = default_color

            color_row_frame = ttk.Frame(self.scrollable_frame)
            color_row_frame.pack(fill="x", padx=5, pady=2)

            label = ttk.Label(color_row_frame, text=f"{product}:")
            label.pack(side="left", padx=5)

            color_canvas = tk.Canvas(color_row_frame, width=20, height=20, bg=default_color, relief="groove", bd=1)
            color_canvas.pack(side="left", padx=5)
            self.category_color_frames[product] = color_canvas

            choose_button = ttk.Button(color_row_frame, text="色を選択",
                                       command=lambda p=product: self.choose_color_for_category(p))
            choose_button.pack(side="left", padx=5)

    def clear_category_color_widgets(self):
        for widget in self.scrollable_frame.winfo_children():
            widget.destroy()
        self.category_colors.clear()
        self.category_color_frames.clear()

    def choose_color_for_category(self, category_name):
        current_color = self.category_colors.get(category_name, 'black')
        color_code = colorchooser.askcolor(title=f"{category_name} の色を選択", initialcolor=current_color)

        if color_code[1]:
            hex_color = color_code[1]
            self.category_colors[category_name] = hex_color
            if category_name in self.category_color_frames:
                self.category_color_frames[category_name].config(bg=hex_color)

    def browse_csv_file(self):
        file_path = filedialog.askopenfilename(
            title="CSVファイルを選択",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        if file_path:
            self.entries["CSVファイルパス:"].delete(0, tk.END)
            self.entries["CSVファイルパス:"].insert(0, file_path)
            self.load_default_colors()

    def browse_save_directory(self):
        dir_path = filedialog.askdirectory(title="保存先フォルダを選択")
        if dir_path:
            self.entries["保存先フォルダ:"].delete(0, tk.END)
            self.entries["保存先フォルダ:"].insert(0, dir_path)

    def generate_chart(self):
        csv_path = self.entries["CSVファイルパス:"].get()
        hanrei_val = self.entries["凡例タイトル:"].get()
        graph_title_val = self.entries["グラフタイトル:"].get()
        xlabel_val = self.entries["X軸ラベル:"].get()
        ylabel_val = self.entries["Y軸ラベル:"].get()
        jiku_val = self.axis_display_var.get()
        save_dir_val = self.entries["保存先フォルダ:"].get()
        save_filename_val = self.entries["保存ファイル名:"].get()
        font_selection = self.font_selection_var.get()
        
        if not csv_path:
            messagebox.showerror("入力エラー", "CSVファイルパスを入力してください。")
            self.status_label.config(text="エラー: CSVファイルパスが未入力です。")
            return
        if not os.path.exists(csv_path):
            messagebox.showerror("入力エラー", f"CSVファイル '{csv_path}' が見つかりません。")
            self.status_label.config(text=f"エラー: CSVファイル '{csv_path}' が見つかりません。")
            return
        if not save_dir_val:
            messagebox.showerror("入力エラー", "保存先フォルダを入力してください。")
            self.status_label.config(text="エラー: 保存先フォルダが未入力です。")
            return
        if not os.path.isdir(save_dir_val):
            try:
                os.makedirs(save_dir_val, exist_ok=True)
            except OSError as e:
                messagebox.showerror("エラー", f"保存先フォルダ '{save_dir_val}' を作成できませんでした: {e}")
                self.status_label.config(text=f"エラー: 保存先フォルダ作成失敗: {e}")
                return
        if not save_filename_val:
            messagebox.showerror("入力エラー", "保存ファイル名を入力してください。")
            self.status_label.config(text="エラー: 保存ファイル名が未入力です。")
            return
        
        if not self.category_colors:
            messagebox.showwarning("警告", "カテゴリの色が設定されていません。CSVを読み込み、カテゴリを検出してください。")

        self.status_label.config(text="グラフを生成中です...")
        self.update_idletasks()

        success, message = create_mekko_chart(csv_path, hanrei_val, graph_title_val,
                                            xlabel_val, ylabel_val, jiku_val,
                                            save_dir_val, save_filename_val, font_selection,
                                            self.category_colors)
        
        if success:
            messagebox.showinfo("成功", message)
            self.status_label.config(text=message, foreground="green")
        else:
            messagebox.showerror("エラー", message)
            self.status_label.config(text=message, foreground="red")

# GUIアプリケーションの起動
if __name__ == "__main__":
    import matplotlib.colors as mcolors
    app = MekkoChartGenerator()
    app.mainloop()