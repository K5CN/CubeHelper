from math import log
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
import pyautogui
import pytesseract
from PIL import Image, ImageTk
import threading
import time
import json
import os
import keyboard # Added for global hotkey

# Configure Tesseract path if necessary (usually not needed if tesseract is in PATH)
# pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe' # Example path

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("魔方助手-by k3")
        self.root.geometry("800x800")
        try:
            # Set application icon
            icon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "icon.png")
            if os.path.exists(icon_path):
                self.root.iconphoto(True, tk.PhotoImage(file=icon_path))
            else:
                print(f"Icon file not found at {icon_path}") # Or log to GUI if preferred
        except Exception as e:
            print(f"Error setting icon: {e}") # Or log to GUI
        # --- UI Elements -

        # Variables
        self.ocr_region = None
        self.refresh_region = None
        self.rules = []
        self.selected_rules_indices = []
        self.is_running = False
        self.thread = None

        # --- UI Elements ---
        # Frame for region selection
        region_frame = tk.LabelFrame(root, text="区域选择", padx=10, pady=10)
        region_frame.pack(padx=10, pady=10, fill="x")

        self.btn_select_ocr = tk.Button(region_frame, text="选择识别区域", command=self.select_ocr_region)
        self.btn_select_ocr.grid(row=0, column=0, padx=5, pady=5)
        self.lbl_ocr_region = tk.Label(region_frame, text="识别区域: 未选择")
        self.lbl_ocr_region.grid(row=0, column=1, padx=5, pady=5, sticky="w")

        self.btn_select_refresh = tk.Button(region_frame, text="选择刷新区域", command=self.select_refresh_region)
        self.btn_select_refresh.grid(row=1, column=0, padx=5, pady=5)
        self.lbl_refresh_region = tk.Label(region_frame, text="刷新区域: 未选择")
        self.lbl_refresh_region.grid(row=1, column=1, padx=5, pady=5, sticky="w")

        # Frame for rules management
        rules_frame = tk.LabelFrame(root, text="规则管理", padx=10, pady=10)
        rules_frame.pack(padx=10, pady=10, fill="x")

        tk.Label(rules_frame, text="规则名称:").grid(row=0, column=0, padx=5, pady=2, sticky="e")
        self.entry_rule_name = tk.Entry(rules_frame, width=20)
        self.entry_rule_name.grid(row=0, column=1, padx=5, pady=2, sticky="w")

        tk.Label(rules_frame, text="词条1:").grid(row=1, column=0, padx=5, pady=2, sticky="e")
        self.entry_attr1 = tk.Entry(rules_frame, width=30)
        self.entry_attr1.grid(row=1, column=1, padx=5, pady=2, sticky="w")

        tk.Label(rules_frame, text="词条2:").grid(row=2, column=0, padx=5, pady=2, sticky="e")
        self.entry_attr2 = tk.Entry(rules_frame, width=30)
        self.entry_attr2.grid(row=2, column=1, padx=5, pady=2, sticky="w")

        tk.Label(rules_frame, text="词条3:").grid(row=3, column=0, padx=5, pady=2, sticky="e")
        self.entry_attr3 = tk.Entry(rules_frame, width=30)
        self.entry_attr3.grid(row=3, column=1, padx=5, pady=2, sticky="w")

        self.btn_add_rule = tk.Button(rules_frame, text="添加规则", command=self.add_rule)
        self.btn_add_rule.grid(row=4, column=0, padx=5, pady=5)
        self.btn_delete_rule = tk.Button(rules_frame, text="删除选中规则", command=self.delete_rule)
        self.btn_delete_rule.grid(row=4, column=1, padx=5, pady=5, sticky="w")

        self.btn_save_rules = tk.Button(rules_frame, text="保存规则", command=self.save_rules)
        self.btn_save_rules.grid(row=5, column=0, padx=5, pady=5)
        self.btn_load_rules = tk.Button(rules_frame, text="加载规则", command=self.load_rules)
        self.btn_load_rules.grid(row=5, column=1, padx=5, pady=5, sticky="w")

        tk.Label(rules_frame, text="已定义规则 (可多选):").grid(row=0, column=2, padx=15, pady=2, sticky="w")
        self.rules_listbox = tk.Listbox(rules_frame, selectmode=tk.MULTIPLE, width=40, height=6)
        self.rules_listbox.grid(row=1, column=2, rowspan=5, padx=15, pady=5, sticky="nsew")
        rules_frame.grid_columnconfigure(2, weight=1)

        # Frame for controls and status
        control_frame = tk.LabelFrame(root, text="控制与状态", padx=10, pady=10)
        control_frame.pack(padx=10, pady=10, fill="x")

        self.btn_start = tk.Button(control_frame, text="开始刷新", command=self.start_process, width=15, height=2)
        self.btn_start.pack(side=tk.LEFT, padx=10, pady=10)
        self.btn_stop = tk.Button(control_frame, text="停止刷新", command=self.stop_process, width=15, height=2, state=tk.DISABLED)
        self.btn_stop.pack(side=tk.LEFT, padx=10, pady=10)

        self.status_label = tk.Label(control_frame, text="状态: 空闲", fg="blue")
        self.status_label.pack(side=tk.LEFT, padx=20, pady=10)

        # Frame for OCR results
        ocr_results_frame = tk.LabelFrame(root, text="OCR识别结果", padx=10, pady=10)
        ocr_results_frame.pack(padx=10, pady=10, fill="both", expand=True)

        self.ocr_text_area = tk.Text(ocr_results_frame, height=10, width=70, state=tk.DISABLED, wrap=tk.WORD)
        self.ocr_text_area_scrollbar = tk.Scrollbar(ocr_results_frame, command=self.ocr_text_area.yview)
        self.ocr_text_area.config(yscrollcommand=self.ocr_text_area_scrollbar.set)
        self.ocr_text_area_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.ocr_text_area.pack(padx=5, pady=5, fill="both", expand=True)

        self.load_rules_from_file() # Load rules on startup if file exists
        # self.root.bind('<Escape>', lambda event: self.stop_process() if self.is_running else None) # Removed, will use global hotkey F12

    def _select_region(self, region_type):
        self.root.withdraw() # Hide main window
        time.sleep(0.2) # Give time for window to hide

        selector_win = tk.Toplevel(self.root)
        selector_win.attributes('-fullscreen', True)
        selector_win.attributes('-alpha', 0.3) # Make it semi-transparent
        selector_win.attributes('-topmost', True)
        selector_win.wait_visibility(selector_win)

        canvas = tk.Canvas(selector_win, cursor="cross", bg="grey")
        canvas.pack(fill=tk.BOTH, expand=True)

        rect_coords = [None, None, None, None]
        rect_id = None

        def on_mouse_press(event):
            rect_coords[0], rect_coords[1] = event.x, event.y
            nonlocal rect_id
            if rect_id:
                canvas.delete(rect_id)
            rect_id = canvas.create_rectangle(0,0,0,0, outline='red', width=2)

        def on_mouse_drag(event):
            rect_coords[2], rect_coords[3] = event.x, event.y
            canvas.coords(rect_id, rect_coords[0], rect_coords[1], rect_coords[2], rect_coords[3])

        def on_mouse_release(event):
            selector_win.destroy()
            self.root.deiconify() # Show main window again

            x1, y1, x2, y2 = min(rect_coords[0], rect_coords[2]), min(rect_coords[1], rect_coords[3]), \
                             max(rect_coords[0], rect_coords[2]), max(rect_coords[1], rect_coords[3])
            
            if x1 is None or y1 is None or x2 is None or y2 is None or x1==x2 or y1==y2:
                messagebox.showerror("错误", "无效的区域选择。")
                return

            selected_region = (x1, y1, x2-x1, y2-y1) # (x, y, width, height)

            if region_type == "ocr":
                self.ocr_region = selected_region
                self.lbl_ocr_region.config(text=f"识别区域: {self.ocr_region}")
            elif region_type == "refresh":
                self.refresh_region = selected_region
                self.lbl_refresh_region.config(text=f"刷新区域: {self.refresh_region} (中心点: {self.refresh_region[0] + self.refresh_region[2]//2}, {self.refresh_region[1] + self.refresh_region[3]//2})")

        canvas.bind("<ButtonPress-1>", on_mouse_press)
        canvas.bind("<B1-Motion>", on_mouse_drag)
        canvas.bind("<ButtonRelease-1>", on_mouse_release)
        selector_win.bind("<Escape>", lambda e: (selector_win.destroy(), self.root.deiconify()))

    def select_ocr_region(self):
        self.status_label.config(text="状态: 正在选择识别区域...", fg="orange")
        self._select_region("ocr")
        self.status_label.config(text="状态: 空闲", fg="blue")

    def select_refresh_region(self):
        self.status_label.config(text="状态: 正在选择刷新区域...", fg="orange")
        self._select_region("refresh")
        self.status_label.config(text="状态: 空闲", fg="blue")

    def add_rule(self):
        name = self.entry_rule_name.get().strip()
        attr1 = self.entry_attr1.get().strip()
        attr2 = self.entry_attr2.get().strip()
        attr3 = self.entry_attr3.get().strip()

        if not name:
            messagebox.showerror("错误", "规则名称不能为空。")
            return
        if not (attr1 or attr2 or attr3):
            messagebox.showerror("错误", "至少需要一个词条。")
            return
        
        # Ensure no duplicate rule names
        for rule in self.rules:
            if rule['name'] == name:
                messagebox.showerror("错误", f"规则名称 '{name}' 已存在。")
                return

        rule_attrs = [a for a in [attr1, attr2, attr3] if a] # Collect non-empty attributes
        if len(rule_attrs) != 3:
             messagebox.showwarning("提示", "规则通常包含3个词条，您输入的词条数量不足3个。 '*' 将被用于填充。")
             while len(rule_attrs) < 3:
                 rule_attrs.append("*")

        self.rules.append({"name": name, "attrs": rule_attrs})
        self.update_rules_listbox()
        self.entry_rule_name.delete(0, tk.END)
        self.entry_attr1.delete(0, tk.END)
        self.entry_attr2.delete(0, tk.END)
        self.entry_attr3.delete(0, tk.END)
        self.status_label.config(text=f"状态: 规则 '{name}' 已添加。", fg="green")

    def delete_rule(self):
        selected_indices = self.rules_listbox.curselection()
        if not selected_indices:
            messagebox.showwarning("提示", "请先选择要删除的规则。")
            return
        
        # Iterate in reverse to avoid index issues when deleting
        for i in sorted(selected_indices, reverse=True):
            deleted_rule_name = self.rules[i]['name']
            del self.rules[i]
            self.status_label.config(text=f"状态: 规则 '{deleted_rule_name}' 已删除。", fg="orange")
        self.update_rules_listbox()

    def update_rules_listbox(self):
        self.rules_listbox.delete(0, tk.END)
        for rule in self.rules:
            self.rules_listbox.insert(tk.END, f"{rule['name']}: {', '.join(rule['attrs'])}")

    def get_rules_filepath(self):
        return os.path.join(os.path.dirname(os.path.abspath(__file__)), "rules.json")

    def save_rules(self):
        if not self.rules:
            messagebox.showinfo("提示", "没有规则可以保存。")
            return
        filepath = self.get_rules_filepath()
        try:
            with open(filepath, 'w', encoding='utf-8') as f:
                json.dump(self.rules, f, ensure_ascii=False, indent=4)
            messagebox.showinfo("成功", f"规则已保存到 {filepath}")
            self.status_label.config(text="状态: 规则已保存。", fg="green")
        except Exception as e:
            messagebox.showerror("错误", f"保存规则失败: {e}")
            self.status_label.config(text="状态: 保存规则失败。", fg="red")

    def load_rules_from_file(self):
        filepath = self.get_rules_filepath()
        if os.path.exists(filepath):
            try:
                with open(filepath, 'r', encoding='utf-8') as f:
                    self.rules = json.load(f)
                self.update_rules_listbox()
                self.status_label.config(text="状态: 规则已从文件加载。", fg="blue")
            except Exception as e:
                messagebox.showerror("错误", f"加载规则失败: {e}")
                self.rules = [] # Reset rules if loading fails
                self.update_rules_listbox()
                self.status_label.config(text="状态: 加载规则失败。", fg="red")
        else:
            self.status_label.config(text="状态: 未找到规则文件，未加载任何规则。", fg="orange")
            
    def load_rules(self):
        self.load_rules_from_file()
        if self.rules:
             messagebox.showinfo("成功", "规则已加载。")
        elif not os.path.exists(self.get_rules_filepath()):
            messagebox.showinfo("提示", "未找到规则文件。")

    def perform_ocr(self):
        if not self.ocr_region:
            return None
        try:
            # Ensure region is valid
            x, y, w, h = self.ocr_region
            if w <= 0 or h <= 0:
                self.log_to_gui("OCR错误: 识别区域无效 (宽度或高度为0或负数)。")
                return None

            screenshot = pyautogui.screenshot(region=self.ocr_region)
            # screenshot.save("debug_ocr_capture.png") # For debugging
            # Preprocessing can be added here if needed (e.g., grayscale, thresholding)
            # For example:
            # gray_image = screenshot.convert('L')
            # threshold_image = gray_image.point(lambda p: p > 128 and 255)
            text = pytesseract.image_to_string(screenshot, lang='eng+chi_sim') # Adjust lang as needed
            self.log_to_gui(f"识别到的文本:\n{text}")
            return text.strip().split('\n')
        except Exception as e:
            self.log_to_gui(f"OCR错误: {e}")
            # self.log_to_gui(f"OCR区域: {self.ocr_region}") # Log region for debugging
            return None

    # Note: The log_to_gui method already exists and seems functional with timestamps and colors.
    # The request was to ensure OCR results are displayed, which perform_ocr already does by calling log_to_gui.
    # No change needed for log_to_gui itself based on the new file content, but its usage in process_loop will be key.

    def check_rules(self, lines):
        if not lines or not self.selected_rules_indices:
            return False
        
        cleaned_lines = [line.strip() for line in lines if line.strip()] # Remove empty lines and strip
        if len(cleaned_lines) < 3: # Expecting at least 3 lines for attributes
            self.log_to_gui(f"规则检查: 识别到的有效行数 ({len(cleaned_lines)}) 少于3行。")
            return False

        # Consider only the first 3 significant lines for matching, or more if needed
        # For now, let's assume the first 3 lines are the most relevant attributes.
        # This might need adjustment based on actual OCR output variance.
        # Example: if OCR sometimes picks up "Legendary" or other text before attributes.
        # A more robust approach would be to filter lines that look like attributes.
        
        # Simple approach: take the first 3 non-empty lines
        # A more robust approach might involve looking for lines with typical attribute patterns (e.g., containing '%', '+', ':')
        # For now, we'll use the first 3 lines if available.
        # If less than 3 lines, it won't match a 3-attribute rule.
        
        # Let's try to be a bit smarter: filter lines that are likely attributes
        potential_attr_lines = []
        for line in cleaned_lines:
            # Heuristic: attributes often contain ':', '%', or '+'
            # Also, they are not too long and not just numbers (like ATT Increase)
            if (':' in line or '%' in line or '+' in line) and not line.lower().startswith(("att increase", "combat power", "legendary")):
                potential_attr_lines.append(line)
        
        if len(potential_attr_lines) < 3:
            self.log_to_gui(f"规则检查: 识别到的潜在词条行数 ({len(potential_attr_lines)}) 少于3行。 原始行: {cleaned_lines}")
            return False

        # Use the first 3 identified potential attribute lines
        current_attrs_text = potential_attr_lines[:3]
        self.log_to_gui(f"用于匹配的词条: {current_attrs_text}")

        for rule_idx in self.selected_rules_indices:
            rule = self.rules[rule_idx]
            rule_name = rule['name']
            target_attrs = rule['attrs']
            self.log_to_gui(f"正在检查规则 '{rule_name}': {target_attrs}")

            # Exact match for specified parts, wildcard for '*'
            # Check if all target_attrs can be found in current_attrs_text (order-agnostic)
            
            temp_current_attrs = list(current_attrs_text) # Make a copy to "consume" matched lines
            matches_found = 0

            for target_attr_full in target_attrs:
                target_attr_key = target_attr_full.split(':')[0].strip() if ':' in target_attr_full else target_attr_full.strip()
                target_attr_val_part = target_attr_full.split(':',1)[1].strip() if ':' in target_attr_full else ""
                
                if target_attr_key == "*":
                    matches_found += 1
                    # Try to remove a line from temp_current_attrs so it's not matched again by another '*' or specific rule
                    # This is tricky. For now, a '*' matches if there's any line left.
                    # A better way for '*' might be to ensure it doesn't prevent specific matches.
                    # Let's assume '*' consumes one available line if not already matched by a specific part of the rule.
                    # This logic needs refinement for complex '*' interactions.
                    # For simplicity now: if a rule has '*', it means that slot can be anything.
                    continue # Wildcard matches anything for its slot

                found_this_target_attr = False
                for i, current_attr_full in enumerate(temp_current_attrs):
                    current_attr_key = current_attr_full.split(':')[0].strip() if ':' in current_attr_full else current_attr_full.strip()
                    current_attr_val_part = current_attr_full.split(':',1)[1].strip() if ':' in current_attr_full else ""
                    
                    # Rule: "ATT:" in "Magic ATT:" is not a match.
                    # So, current_attr_key must be exactly target_attr_key
                    if current_attr_key == target_attr_key:
                        # If target_attr_val_part is specified, it must also match
                        if target_attr_val_part:
                            if target_attr_val_part == current_attr_val_part:
                                matches_found += 1
                                found_this_target_attr = True
                                temp_current_attrs.pop(i) # Consume this line
                                break
                        else: # No value specified in rule, so key match is enough
                            matches_found += 1
                            found_this_target_attr = True
                            temp_current_attrs.pop(i) # Consume this line
                            break
                if not found_this_target_attr and target_attr_key != "*": # If a specific attribute wasn't found
                    self.log_to_gui(f"规则 '{rule_name}' 未匹配: 词条 '{target_attr_full}' 未在识别结果中找到或值不匹配。")
                    break # This rule doesn't match
            
            if matches_found == len(target_attrs): # All parts of the rule matched
                self.log_to_gui(f"成功! 规则 '{rule_name}' 已匹配。", "green")
                messagebox.showinfo("成功", f"找到匹配词条! 规则: {rule_name}\n词条:\n{chr(10).join(current_attrs_text)}")
                return True
        return False

    def refresh_action(self):
        if not self.refresh_region:
            self.log_to_gui("刷新错误: 未选择刷新区域。", "red")
            return
        try:
            # Click center of refresh region
            center_x = self.refresh_region[0] + self.refresh_region[2] // 2
            center_y = self.refresh_region[1] + self.refresh_region[3] // 2
            pyautogui.click(center_x, center_y)
            self.log_to_gui(f"点击刷新区域中心: ({center_x}, {center_y})")
            time.sleep(0.1) # Small delay after click
            pyautogui.press('enter')
            time.sleep(0.05) # Small interval between presses
            pyautogui.press('enter')
            time.sleep(0.05)
            pyautogui.press('enter')
            self.log_to_gui("已按三次Enter键。")
            time.sleep(0.5) # Wait for game UI to update, adjust if necessary
        except Exception as e:
            self.log_to_gui(f"刷新操作失败: {e}", "red")

    def process_loop(self):
        attempts = 0
        while self.is_running:
            attempts += 1
            self.log_to_gui(f"--- 第 {attempts} 次尝试 ---")
            
            # 1. Perform OCR
            ocr_lines = self.perform_ocr()
            if not self.is_running: break # Check if stopped during OCR
            if not ocr_lines:
                self.log_to_gui("OCR未能识别到文本，尝试刷新后重试...", "orange")
                if self.is_running: self.refresh_action() # Try to refresh if OCR fails
                if not self.is_running: break
                time.sleep(1) # Wait after refresh before next OCR
                continue

            # 2. Check rules
            if self.is_running and self.check_rules(ocr_lines): # check_rules now shows messagebox and returns True/False
                # If check_rules returns True, it means a match was found.
                # The messagebox.showinfo in check_rules will alert the user.
                # We need to stop the process here.
                self.log_to_gui("匹配成功，准备停止。", "green")
                self.is_running = False # Signal loop to stop
                # GUI updates will be handled by _finalize_stop_state called after loop exits
            
            if not self.is_running: break # Check if stopped (either by match or by user)

            # 3. Refresh if no match and still running
            if self.is_running:
                self.log_to_gui("未匹配到规则，执行刷新操作...")
                self.refresh_action()
            
            if not self.is_running: break # Check if stopped during/after refresh

            # time.sleep(0.1) # Reduced delay, OCR and refresh have sleeps
        
        # This part runs after the loop finishes
        self.root.after(0, self._finalize_stop_state) # Ensure GUI updates are on main thread

    def _finalize_stop_state(self):
        """Helper to update GUI elements after processing stops. Called from main thread."""
        self.btn_start.config(state=tk.NORMAL)
        self.btn_stop.config(state=tk.DISABLED)
        
        # Determine final status message based on how it stopped
        # This requires more sophisticated state tracking if we want to distinguish
        # between manual stop and match found, beyond just log messages.
        # For now, check_rules logs success clearly.
        ocr_content_snapshot = self.ocr_text_area.get("1.0", tk.END) # Get current content
        if "成功! 规则" in ocr_content_snapshot and "已匹配" in ocr_content_snapshot:
            final_message = "状态: 匹配成功并停止!"
            final_color = "green"
            self.log_to_gui("匹配成功，已停止!", "green") # Log final confirmation
        elif "用户请求停止" in ocr_content_snapshot or "已手动停止" in ocr_content_snapshot : # Check if stop was manual
            final_message = "状态: 已手动停止"
            final_color = "orange"
        else: # Generic stop
            final_message = "状态: 已停止"
            final_color = "blue"
            if not ("匹配成功" in ocr_content_snapshot or "已手动停止" in ocr_content_snapshot):
                 self.log_to_gui("刷新已停止。", "blue")

        self.status_label.config(text=final_message, fg=final_color)
        self.thread = None # Clear thread reference
        try:
            keyboard.remove_hotkey('f12')
            self.log_to_gui("F12 热键已注销。", "blue")
        except Exception as e:
            # self.log_to_gui(f"注销 F12 热键时出错: {e}", "yellow") # Minor, might not have been registered
            pass

    def start_process(self):
        if self.is_running:
            messagebox.showwarning("提示", "刷新已经在运行中。")
            return

        if not self.ocr_region:
            messagebox.showerror("错误", "请先选择识别区域。")
            return
        if not self.refresh_region:
            messagebox.showerror("错误", "请先选择刷新区域。")
            return
        
        self.selected_rules_indices = self.rules_listbox.curselection()
        if not self.selected_rules_indices:
            messagebox.showerror("错误", "请至少选择一个规则来匹配。")
            return
        
        self.log_to_gui("所选规则:", clear_before_log=True)
        for idx in self.selected_rules_indices:
            self.log_to_gui(f"  - {self.rules[idx]['name']}: {self.rules[idx]['attrs']}")

        self.is_running = True
        self.btn_start.config(state=tk.DISABLED)
        self.btn_stop.config(state=tk.NORMAL)
        self.status_label.config(text="状态: 运行中...", fg="green")
        self.log_to_gui("开始刷新过程... 按 F12 停止", "blue")
        try:
            keyboard.add_hotkey('f12', self.stop_process_from_hotkey, suppress=True)
            self.log_to_gui("F12 热键已注册用于停止。", "blue")
        except Exception as e:
            self.log_to_gui(f"注册 F12 热键失败: {e}。请尝试以管理员权限运行程序。", "red")

        self.thread = threading.Thread(target=self.process_loop, daemon=True)
        self.thread.start()

    def stop_process_from_hotkey(self):
        if self.is_running:
            # This is called from the keyboard listener thread.
            # We need to schedule GUI updates and critical logic on the main Tkinter thread.
            self.root.after(0, self._initiate_stop_from_hotkey)

    def _initiate_stop_from_hotkey(self):
        # This method runs in the main Tkinter thread.
        if self.is_running: # Double check state in case it changed
            self.is_running = False # Signal the loop to stop
            self.log_to_gui("F12 热键按下，正在停止...", "orange")
            # The process_loop will see self.is_running as False and then call _finalize_stop_state
            # Hotkey removal will be handled in _finalize_stop_state

    def stop_process(self):
        if self.is_running:
            self.is_running = False # Signal the loop to stop
            self.log_to_gui("停止按钮按下，正在停止...", "orange")
            # The process_loop will see self.is_running as False and then call _finalize_stop_state
            # Hotkey removal will be handled in _finalize_stop_state
        else:
            # If already stopped or not started, ensure UI is in a consistent idle state.
            if not self.thread: # only update if not in the middle of stopping already
                self.btn_start.config(state=tk.NORMAL)
                self.btn_stop.config(state=tk.DISABLED)
                self.status_label.config(text="状态: 空闲", fg="blue")
                self.log_to_gui("当前没有刷新任务在运行或已停止。", "blue")
                # Ensure hotkey is removed if stop is pressed when not running
                try:
                    keyboard.remove_hotkey('f12')
                    # self.log_to_gui("F12 热键已尝试注销 (如果之前已注册)。", "blue")
                except Exception: 
                    self.log_to_gui("停止失败: F12 热键未能注销。", "red")
                    pass

    def log_to_gui(self, message, color="black", clear_before_log=False):
        self.ocr_text_area.config(state=tk.NORMAL)
        if clear_before_log:
            self.ocr_text_area.delete('1.0', tk.END)
        
        # Add timestamp
        timestamp = time.strftime("%H:%M:%S", time.localtime())
        formatted_message = f"[{timestamp}] {message}\n"

        # Create a tag for color if it doesn't exist
        tag_name = f"color_{color}"
        if not tag_name in self.ocr_text_area.tag_names():
            self.ocr_text_area.tag_configure(tag_name, foreground=color)
        
        self.ocr_text_area.insert(tk.END, formatted_message, tag_name)
        self.ocr_text_area.see(tk.END) # Scroll to the end
        self.ocr_text_area.config(state=tk.DISABLED)
        self.root.update_idletasks() # Ensure GUI updates

    def on_closing(self):
        if self.is_running:
            if messagebox.askyesno("退出", "刷新正在进行中，确定要退出吗？"):
                self.stop_process()
                self.root.destroy()
            else:
                return # Do not close
        else:
            self.root.destroy()

if __name__ == "__main__":
    # Check for Tesseract installation (optional, basic check)
    try:
        pytesseract.get_tesseract_version()
    except pytesseract.TesseractNotFoundError:
        messagebox.showerror("Tesseract 未找到", 
                             "Tesseract OCR 未安装或未在系统PATH中。\n"+
                             "请确保已安装Tesseract并将其添加到环境变量PATH中。\n"+
                             "您可能需要在代码中手动设置 `pytesseract.pytesseract.tesseract_cmd`。")
        # exit() # Decide if you want to exit or let user try to set path
    except Exception as e:
        messagebox.showwarning("Tesseract警告", f"检查Tesseract时发生错误: {e}")

    main_root = tk.Tk()
    app = App(main_root)
    main_root.protocol("WM_DELETE_WINDOW", app.on_closing) # Handle window close button
    main_root.mainloop()