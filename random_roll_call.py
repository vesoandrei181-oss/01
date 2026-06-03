import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
import random
import os
from datetime import datetime
from openpyxl import Workbook, load_workbook

STUDENTS_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "students.txt")
EXCEL_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "records.xlsx")
GRID_COLS = 5
CARD_FONT = ("Microsoft YaHei", 12)
PICKED_FONT = ("Microsoft YaHei", 28, "bold")
HIGHLIGHT_BG = "#FF6B6B"
HIGHLIGHT_FG = "#FFFFFF"
NORMAL_BG = "#E8F4FD"
NORMAL_FG = "#333333"
BUTTON_BG = "#4ECDC4"
BUTTON_FG = "#FFFFFF"


class RandomRollCallApp:
    def __init__(self, root):
        self.root = root
        self.root.title("随机点名系统")
        self.root.geometry("900x700")
        self.root.configure(bg="#F7F9FC")
        self.root.minsize(700, 500)

        self.students = []
        self.records = []
        self.name_labels = {}
        self.picked_name = None
        self.picked_label_ref = None
        self._rolling = False
        self._roll_job = None

        self._build_ui()
        self._load_students()

    def _build_ui(self):
        # Title
        title_frame = tk.Frame(self.root, bg="#F7F9FC")
        title_frame.pack(pady=(15, 5))
        tk.Label(
            title_frame, text="🎯 随机点名系统",
            font=("Microsoft YaHei", 22, "bold"), fg="#2C3E50", bg="#F7F9FC"
        ).pack()

        # Name cards area (scrollable)
        canvas_frame = tk.Frame(self.root, bg="#F7F9FC")
        canvas_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=5)

        self.canvas = tk.Canvas(canvas_frame, bg="#F7F9FC", highlightthickness=0)
        scrollbar = tk.Scrollbar(canvas_frame, orient=tk.VERTICAL, command=self.canvas.yview)
        self.name_grid_frame = tk.Frame(self.canvas, bg="#F7F9FC")

        self.name_grid_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )

        self.canvas_window = self.canvas.create_window((0, 0), window=self.name_grid_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=scrollbar.set)

        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Make canvas resize with window
        canvas_frame.bind("<Configure>", self._on_canvas_resize)
        self.root.bind("<MouseWheel>", self._on_mousewheel)

        # Picked name display
        self.picked_frame = tk.Frame(self.root, bg="#FFF3CD", relief=tk.RIDGE, bd=3)
        self.picked_frame.pack(fill=tk.X, padx=20, pady=5)
        self.picked_display = tk.Label(
            self.picked_frame,
            text="点击「随机点名」开始",
            font=("Microsoft YaHei", 18, "bold"),
            fg="#856404", bg="#FFF3CD"
        )
        self.picked_display.pack(pady=12)

        # Status bar
        self.status_label = tk.Label(
            self.root, text="已加载 0 名学生 | 当前记录: 0 条",
            font=("Microsoft YaHei", 10), fg="#6C757D", bg="#F7F9FC"
        )
        self.status_label.pack(pady=(0, 5))

        # Buttons
        btn_frame = tk.Frame(self.root, bg="#F7F9FC")
        btn_frame.pack(pady=(0, 15))

        self._create_btn(btn_frame, "🎲 随机点名", self._pick_random, "#FF6B6B").pack(
            side=tk.LEFT, padx=8, ipadx=20, ipady=6
        )
        self._create_btn(btn_frame, "📂 导入名单", self._import_students, "#4ECDC4").pack(
            side=tk.LEFT, padx=8, ipadx=20, ipady=6
        )
        self._create_btn(btn_frame, "💾 保存到Excel", self._save_to_excel, "#45B7D1").pack(
            side=tk.LEFT, padx=8, ipadx=20, ipady=6
        )
        self._create_btn(btn_frame, "📋 查看记录", self._view_records, "#96CEB4").pack(
            side=tk.LEFT, padx=8, ipadx=20, ipady=6
        )

    def _create_btn(self, parent, text, command, color):
        return tk.Button(
            parent, text=text, command=command,
            font=("Microsoft YaHei", 12, "bold"),
            bg=color, fg="white", activebackground=color,
            relief=tk.FLAT, cursor="hand2",
            padx=10, pady=4
        )

    def _on_canvas_resize(self, event):
        canvas_width = event.width
        self.canvas.itemconfig(self.canvas_window, width=canvas_width)

    def _on_mousewheel(self, event):
        self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def _load_students(self):
        if os.path.exists(STUDENTS_FILE):
            with open(STUDENTS_FILE, "r", encoding="utf-8") as f:
                self.students = [line.strip() for line in f if line.strip()]
        if not self.students:
            self.students = ["示例学生1", "示例学生2", "示例学生3"]
        self._refresh_name_grid()
        self._update_status()

    def _refresh_name_grid(self):
        for widget in self.name_grid_frame.winfo_children():
            widget.destroy()
        self.name_labels.clear()
        self.picked_label_ref = None

        for i, name in enumerate(self.students):
            row = i // GRID_COLS
            col = i % GRID_COLS

            card = tk.Frame(
                self.name_grid_frame, bg=NORMAL_BG,
                relief=tk.RAISED, bd=2,
                highlightthickness=1, highlightbackground="#DEE2E6"
            )
            card.grid(row=row, column=col, padx=6, pady=6, sticky="nsew")
            card.grid_propagate(False)
            card.config(width=150, height=50)

            label = tk.Label(
                card, text=name, font=CARD_FONT,
                bg=NORMAL_BG, fg=NORMAL_FG
            )
            label.place(relx=0.5, rely=0.5, anchor="center")
            label.bind("<Button-1>", lambda e, n=name: self._manual_pick(n))

            self.name_labels[name] = (card, label)

        # Configure grid columns
        for col in range(GRID_COLS):
            self.name_grid_frame.grid_columnconfigure(col, weight=1, uniform="col")

    def _manual_pick(self, name):
        self._highlight_name(name)

    def _highlight_name(self, name):
        # Reset all cards to normal
        for n, (card, label) in self.name_labels.items():
            card.config(bg=NORMAL_BG, highlightbackground="#DEE2E6")
            label.config(bg=NORMAL_BG, fg=NORMAL_FG, font=CARD_FONT)

        # Highlight picked
        if name in self.name_labels:
            card, label = self.name_labels[name]
            card.config(bg=HIGHLIGHT_BG, highlightbackground="#E53935")
            label.config(bg=HIGHLIGHT_BG, fg=HIGHLIGHT_FG, font=("Microsoft YaHei", 16, "bold"))
            self.picked_label_ref = (card, label)

        self.picked_name = name
        self.picked_display.config(
            text=f"⭐ 当前被点中: {name} ⭐",
            fg="#D4A017"
        )

    def _pick_random(self):
        if not self.students:
            messagebox.showwarning("提示", "请先导入学生名单！")
            return
        if self._rolling:
            return

        self._rolling = True
        self._final_name = random.choice(self.students)
        # Total animation steps: ~20 rounds, ~1.5 seconds
        self._roll_steps = 20
        self._roll_current = 0
        self._roll_delay = 30  # start fast
        self._do_roll_step()

    def _do_roll_step(self):
        if self._roll_current >= self._roll_steps:
            # Animation done, settle on final name
            self._rolling = False
            self._highlight_name(self._final_name)
            self._ask_answer_and_score()
            return

        # Pick a random name to flash during animation
        flash_name = random.choice(self.students)
        self._highlight_name(flash_name)
        # Also update the picked display with a rolling indicator
        self.picked_display.config(
            text=f"🎰 抽取中... {flash_name}",
            fg="#E65100"
        )

        self._roll_current += 1
        # Gradually increase delay to create deceleration effect
        progress = self._roll_current / self._roll_steps
        delay = int(30 + 200 * progress * progress)  # quadratic ease-out

        self._roll_job = self.root.after(delay, self._do_roll_step)

    def _ask_answer_and_score(self):
        name = self._final_name
        answer = simpledialog.askstring(
            "回答记录", f"请记录 {name} 的回答内容：",
            parent=self.root
        )
        if answer is None:
            return
        if not answer.strip():
            messagebox.showwarning("提示", "回答内容不能为空！")
            return

        score = simpledialog.askinteger(
            "评分", f"请为 {name} 的回答评分 (1-10)：",
            parent=self.root, minvalue=1, maxvalue=10
        )
        if score is None:
            return

        record = {
            "name": name,
            "answer": answer.strip(),
            "score": score,
            "time": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        }
        self.records.append(record)
        self._update_status()
        messagebox.showinfo("成功", f"已记录 {name} 的回答，评分: {score}/10")

    def _import_students(self):
        filepath = filedialog.askopenfilename(
            title="选择学生名单文件",
            filetypes=[("文本文件", "*.txt"), ("所有文件", "*.*")]
        )
        if not filepath:
            return

        with open(filepath, "r", encoding="utf-8") as f:
            new_students = [line.strip() for line in f if line.strip()]

        if not new_students:
            messagebox.showwarning("提示", "文件中没有找到有效的学生姓名！")
            return

        self.students = new_students
        self.records.clear()
        self.picked_name = None
        self.picked_display.config(text="点击「随机点名」开始", fg="#856404")
        self._refresh_name_grid()
        self._update_status()
        messagebox.showinfo("成功", f"已导入 {len(self.students)} 名学生")

    def _save_to_excel(self):
        if not self.records:
            messagebox.showwarning("提示", "暂无记录可保存！")
            return

        if os.path.exists(EXCEL_FILE):
            wb = load_workbook(EXCEL_FILE)
            ws = wb.active
            existing_names = set()
            for row in ws.iter_rows(min_row=2, values_only=True):
                if row[1]:
                    existing_names.add(row[1])
            for r in self.records:
                if r["name"] not in existing_names:
                    existing_names.add(r["name"])
        else:
            wb = Workbook()
            ws = wb.active
            ws.title = "点名记录"
            ws.append(["序号", "姓名", "回答内容", "评分", "时间"])
            existing_names = set()

        # Always create a summary sheet
        if "成绩汇总" in wb.sheetnames:
            del wb["成绩汇总"]
        ws_summary = wb.create_sheet("成绩汇总")
        ws_summary.append(["姓名", "被点名次数", "总评分", "平均评分", "回答记录汇总"])

        # Rebuild records sheet
        ws.delete_rows(1, ws.max_row)
        ws.append(["序号", "姓名", "回答内容", "评分", "时间"])
        for i, r in enumerate(self.records, 1):
            ws.append([i, r["name"], r["answer"], r["score"], r["time"]])

        # Build summary
        summary = {}
        for r in self.records:
            name = r["name"]
            if name not in summary:
                summary[name] = {"count": 0, "total": 0, "answers": []}
            summary[name]["count"] += 1
            summary[name]["total"] += r["score"]
            summary[name]["answers"].append(f"[{r['score']}分] {r['answer']}")

        # Also include students with no records
        for name in self.students:
            if name not in summary:
                summary[name] = {"count": 0, "total": 0, "answers": []}

        for name, data in summary.items():
            avg = round(data["total"] / data["count"], 1) if data["count"] > 0 else 0
            answers_text = "; ".join(data["answers"]) if data["answers"] else "无记录"
            ws_summary.append([name, data["count"], data["total"], avg, answers_text])

        # Adjust column widths
        ws.column_dimensions["A"].width = 6
        ws.column_dimensions["B"].width = 12
        ws.column_dimensions["C"].width = 50
        ws.column_dimensions["D"].width = 8
        ws.column_dimensions["E"].width = 20

        ws_summary.column_dimensions["A"].width = 12
        ws_summary.column_dimensions["B"].width = 12
        ws_summary.column_dimensions["C"].width = 10
        ws_summary.column_dimensions["D"].width = 10
        ws_summary.column_dimensions["E"].width = 60

        wb.save(EXCEL_FILE)
        self._update_status()
        messagebox.showinfo("成功", f"已保存 {len(self.records)} 条记录到 {EXCEL_FILE}")

    def _view_records(self):
        if not self.records:
            messagebox.showinfo("记录", "暂无记录")
            return

        win = tk.Toplevel(self.root)
        win.title("点名记录")
        win.geometry("700x400")
        win.configure(bg="#F7F9FC")

        frame = tk.Frame(win, bg="#F7F9FC")
        frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        scrollbar = tk.Scrollbar(frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        text = tk.Text(
            frame, font=("Microsoft YaHei", 11),
            yscrollcommand=scrollbar.set, wrap=tk.WORD
        )
        text.pack(fill=tk.BOTH, expand=True)
        scrollbar.config(command=text.yview)

        for i, r in enumerate(self.records, 1):
            text.insert(tk.END, f"{i}. {r['name']} | 评分: {r['score']}/10 | {r['time']}\n")
            text.insert(tk.END, f"   回答: {r['answer']}\n")
            text.insert(tk.END, "-" * 50 + "\n")

        text.config(state=tk.DISABLED)

    def _update_status(self):
        self.status_label.config(
            text=f"已加载 {len(self.students)} 名学生 | 当前记录: {len(self.records)} 条"
        )


def main():
    root = tk.Tk()
    app = RandomRollCallApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
