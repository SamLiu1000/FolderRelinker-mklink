# -*- coding: utf-8 -*-
"""
Folder Relinker-mklink
A graphical Windows symbolic link management tool
"""

import sys
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import subprocess
import json
import shutil
from datetime import datetime

class MKLinkManager:
    def __init__(self, root):
        self.root = root
        self.root.title("Folder Relinker-mklink")
        self.root.geometry("1000x600")
        self.root.resizable(True, True)

        # Language setting (default: English)
        self.lang = tk.StringVar(value="en")

        # Multilingual texts
        self.texts = {
            "zh": {
                "title": "Folder Relinker-mklink",
                "operation": "操作区域",
                "link_type": "链接类型",
                "from": "源文件夹",
                "from_browse": "浏览...",
                "to": "目标位置",
                "to_browse": "浏览...",
                "transfer": "转移",
                "type_desc": "链接类型说明",
                "desc_j": "/J - 本地文件夹，不支持网络路径，推荐",
                "desc_d": "/D - 本地文件夹，支持网络路径，可能存在权限和兼容性问题",
                "records": "操作记录",
                "col_type": "类型",
                "col_from": "源路径",
                "col_to": "目标路径",
                "col_name": "文件名",
                "add_record": "添加记录",
                "delete_record": "删除记录",
                "clear_records": "清空记录",
                "restore_selected": "恢复选中",
                "restore_all": "一键恢复",
                "select_from": "选择源文件夹",
                "select_to": "选择目标位置",
                "error_from": "请选择源文件夹",
                "error_to": "请选择目标位置",
                "error_from_exist": "源路径不存在",
                "error_to_exist": "目标位置已存在",
                "error_link": "创建链接失败，请确保以管理员身份运行程序。\n文件已恢复到原位置。",
                "success": "转移完成！\n已创建{}链接\n原位置: {}\n新位置: {}",
                "confirm_delete": "确定要删除选中的记录吗？",
                "confirm_clear": "确定要清空所有记录吗？",
                "confirm_restore": "确定要恢复所有链接吗？\n\n这将删除所有符号链接并将文件夹移回原位置。\n此操作不可撤销！",
                "confirm_restore_single": "确定要恢复此链接吗？\n\n源路径: {}\n目标路径: {}\n\n这将删除符号链接并将文件夹移回原位置。",
                "restore_success": "恢复成功！\n已恢复 {} 个链接",
                "restore_success_single": "恢复成功！\n已恢复链接\n源路径: {}\n目标路径: {}",
                "restore_failed": "恢复失败：\n{}",
                "restore_partial": "部分恢复成功\n成功: {}\n失败: {}",
                "warning_select": "请先选择要删除的记录",
                "warning_select_restore": "请先选择要恢复的记录",
                "error_source_not_link": "源路径不是符号链接或目录",
                "error_target_not_found": "目标路径不存在",
                "error_source_exists": "源路径已存在非链接文件，无法恢复",
                "dialog_title": "添加记录",
                "dialog_link_type": "链接类型:",
                "dialog_from": "源路径 (链接位置):",
                "dialog_to": "目标路径 (实际位置):",
                "dialog_confirm": "确认添加",
                "dialog_error": "请填写完整信息",
            },
            "en": {
                "title": "Folder Relinker-mklink",
                "operation": "Operation",
                "link_type": "Link Type",
                "from": "From",
                "from_browse": "Browse...",
                "to": "To",
                "to_browse": "Browse...",
                "transfer": "Transfer",
                "type_desc": "Link Type Description",
                "desc_j": "/J - Local folders only, no network paths, recommended",
                "desc_d": "/D - Local folders, supports network paths, may have permission/compatibility issues",
                "records": "Operation Records",
                "col_type": "Type",
                "col_from": "From",
                "col_to": "To",
                "col_name": "Name",
                "add_record": "Add Record",
                "delete_record": "Delete Record",
                "clear_records": "Clear All",
                "restore_selected": "Restore Selected",
                "restore_all": "Restore All",
                "select_from": "Select Source Folder",
                "select_to": "Select Target Location",
                "error_from": "Please select source folder",
                "error_to": "Please select target location",
                "error_from_exist": "Source path does not exist",
                "error_to_exist": "Target already exists",
                "error_link": "Failed to create link. Please run as Administrator.\nFile restored to original location.",
                "success": "Transfer complete!\nCreated {} link\nOriginal: {}\nNew location: {}",
                "confirm_delete": "Delete selected record(s)?",
                "confirm_clear": "Clear all records?",
                "confirm_restore": "Are you sure you want to restore all links?\n\nThis will remove all symbolic links and move folders back to original locations.\nThis action cannot be undone!",
                "confirm_restore_single": "Restore this link?\n\nSource: {}\nTarget: {}\n\nThis will remove the symbolic link and move folder back to original location.",
                "restore_success": "Restore successful!\nRestored {} links",
                "restore_success_single": "Restore successful!\nSource: {}\nTarget: {}",
                "restore_failed": "Restore failed:\n{}",
                "restore_partial": "Partial restore\nSuccess: {}\nFailed: {}",
                "warning_select": "Please select record(s) to delete",
                "warning_select_restore": "Please select a record to restore",
                "error_source_not_link": "Source path is not a symbolic link or directory",
                "error_target_not_found": "Target path does not exist",
                "error_source_exists": "Source path already exists as non-link, cannot restore",
                "dialog_title": "Add Record",
                "dialog_link_type": "Link Type:",
                "dialog_from": "Source Path (Link Location):",
                "dialog_to": "Target Path (Actual Location):",
                "dialog_confirm": "Confirm",
                "dialog_error": "Please fill in all fields",
            }
        }

        # Records file path
        # When packaged as EXE, use the directory of the executable
        if getattr(sys, 'frozen', False):
            base_dir = os.path.dirname(sys.executable)
        else:
            base_dir = os.path.dirname(os.path.abspath(__file__))
        self.records_file = os.path.join(base_dir, "records.json")
        self.records = []

        # Load records
        self.load_records()

        # Create UI
        self.create_widgets()

    def t(self, key):
        """Get text in current language"""
        return self.texts[self.lang.get()].get(key, key)

    def create_widgets(self):
        # Top language switch
        lang_frame = ttk.Frame(self.root)
        lang_frame.pack(fill=tk.X, padx=10, pady=5)

        ttk.Label(lang_frame, text="Language:").pack(side=tk.LEFT)
        ttk.Button(lang_frame, text="EN", width=6, command=lambda: self.switch_lang("en")).pack(side=tk.LEFT, padx=2)
        ttk.Button(lang_frame, text="中文", width=6, command=lambda: self.switch_lang("zh")).pack(side=tk.LEFT, padx=2)

        # Main frame
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Configure grid weights for 1:3 ratio
        main_frame.columnconfigure(0, weight=1)  # Left side
        main_frame.columnconfigure(1, weight=3)  # Right side
        main_frame.rowconfigure(0, weight=1)

        # Left operation area
        self.left_frame = ttk.LabelFrame(main_frame, text=self.t("operation"), padding="10")
        self.left_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 5))

        # Link type selection
        self.type_frame = ttk.LabelFrame(self.left_frame, text=self.t("link_type"), padding="10")
        self.type_frame.pack(fill=tk.X, pady=(0, 10))

        self.link_type = tk.StringVar(value="/J")

        self.rb_j = ttk.Radiobutton(self.type_frame, text="/J", variable=self.link_type, value="/J")
        self.rb_j.pack(anchor=tk.W)
        self.rb_d = ttk.Radiobutton(self.type_frame, text="/D", variable=self.link_type, value="/D")
        self.rb_d.pack(anchor=tk.W)

        # Source folder selection
        self.from_frame = ttk.LabelFrame(self.left_frame, text=self.t("from"), padding="10")
        self.from_frame.pack(fill=tk.X, pady=(0, 10))

        self.from_path = tk.StringVar()
        self.from_entry = ttk.Entry(self.from_frame, textvariable=self.from_path, width=40)
        self.from_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        self.from_btn = ttk.Button(self.from_frame, text=self.t("from_browse"), command=self.browse_from)
        self.from_btn.pack(side=tk.RIGHT)

        # Target location selection
        self.to_frame = ttk.LabelFrame(self.left_frame, text=self.t("to"), padding="10")
        self.to_frame.pack(fill=tk.X, pady=(0, 10))

        self.to_path = tk.StringVar()
        self.to_entry = ttk.Entry(self.to_frame, textvariable=self.to_path, width=40)
        self.to_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        self.to_btn = ttk.Button(self.to_frame, text=self.t("to_browse"), command=self.browse_to)
        self.to_btn.pack(side=tk.RIGHT)

        # Transfer button
        self.transfer_btn = ttk.Button(self.left_frame, text=self.t("transfer"), command=self.execute_transfer)
        self.transfer_btn.pack(fill=tk.X, pady=10)

        # Description area
        self.desc_frame = ttk.LabelFrame(self.left_frame, text=self.t("type_desc"), padding="10")
        self.desc_frame.pack(fill=tk.X, pady=(10, 0))

        self.desc_j_label = ttk.Label(self.desc_frame, text=self.t("desc_j"), wraplength=350)
        self.desc_j_label.pack(anchor=tk.W, pady=2)
        self.desc_d_label = ttk.Label(self.desc_frame, text=self.t("desc_d"), wraplength=350)
        self.desc_d_label.pack(anchor=tk.W, pady=2)

        # Right records area (wider, 1:3 ratio)
        self.right_frame = ttk.LabelFrame(main_frame, text=self.t("records"), padding="10")
        self.right_frame.grid(row=0, column=1, sticky="nsew", padx=(5, 0))

        # Create Treeview for table display with tree structure
        list_frame = ttk.Frame(self.right_frame)
        list_frame.pack(fill=tk.BOTH, expand=True)

        # Define columns: time+type header, path detail
        columns = ("label", "path")
        self.tree = ttk.Treeview(list_frame, columns=columns, show="headings", selectmode="extended")

        # Configure columns
        self.tree.heading("label", text=self.t("col_name") + " / " + self.t("col_type"))
        self.tree.heading("path", text=self.t("col_from") + " / " + self.t("col_to"))

        self.tree.column("label", width=180, minwidth=150, stretch=False)
        self.tree.column("path", width=500, minwidth=300, stretch=True)

        # Scrollbars
        y_scroll = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.tree.yview)
        x_scroll = ttk.Scrollbar(list_frame, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.tree.configure(yscrollcommand=y_scroll.set, xscrollcommand=x_scroll.set)

        # Grid layout for tree and scrollbars
        self.tree.grid(row=0, column=0, sticky="nsew")
        y_scroll.grid(row=0, column=1, sticky="ns")
        x_scroll.grid(row=1, column=0, sticky="ew")

        list_frame.columnconfigure(0, weight=1)
        list_frame.rowconfigure(0, weight=1)

        # Button area at bottom right
        btn_frame = ttk.Frame(self.right_frame)
        btn_frame.pack(fill=tk.X, pady=(10, 0))

        self.add_btn = ttk.Button(btn_frame, text=self.t("add_record"), command=self.add_record_dialog)
        self.add_btn.pack(side=tk.LEFT, padx=(0, 5))
        self.del_btn = ttk.Button(btn_frame, text=self.t("delete_record"), command=self.delete_record)
        self.del_btn.pack(side=tk.LEFT, padx=(0, 5))
        self.clear_btn = ttk.Button(btn_frame, text=self.t("clear_records"), command=self.clear_records)
        self.clear_btn.pack(side=tk.LEFT, padx=(0, 5))

        # Restore Selected button at bottom right
        self.restore_selected_btn = ttk.Button(btn_frame, text=self.t("restore_selected"), command=self.restore_selected)
        self.restore_selected_btn.pack(side=tk.RIGHT, padx=(0, 0))

        # Refresh records display
        self.refresh_records()

    def switch_lang(self, lang):
        """Switch language"""
        self.lang.set(lang)
        self.update_ui_text()

    def update_ui_text(self):
        """Update UI text"""
        self.root.title(self.t("title"))

        # Update left side
        self.left_frame.config(text=self.t("operation"))
        self.type_frame.config(text=self.t("link_type"))
        self.from_frame.config(text=self.t("from"))
        self.to_frame.config(text=self.t("to"))

        self.from_btn.config(text=self.t("from_browse"))
        self.to_btn.config(text=self.t("to_browse"))
        self.transfer_btn.config(text=self.t("transfer"))

        self.desc_frame.config(text=self.t("type_desc"))
        self.desc_j_label.config(text=self.t("desc_j"))
        self.desc_d_label.config(text=self.t("desc_d"))

        # Update right side
        self.right_frame.config(text=self.t("records"))

        # Update tree headings
        self.tree.heading("label", text=self.t("col_name") + " / " + self.t("col_type"))
        self.tree.heading("path", text=self.t("col_from") + " / " + self.t("col_to"))

        self.add_btn.config(text=self.t("add_record"))
        self.del_btn.config(text=self.t("delete_record"))
        self.clear_btn.config(text=self.t("clear_records"))
        self.restore_selected_btn.config(text=self.t("restore_selected"))

        # Refresh records to update language
        self.refresh_records()

    def browse_from(self):
        """Browse source folder"""
        path = filedialog.askdirectory(title=self.t("select_from"))
        if path:
            self.from_path.set(path)

    def browse_to(self):
        """Browse target location"""
        path = filedialog.askdirectory(title=self.t("select_to"))
        if path:
            self.to_path.set(path)

    def execute_transfer(self):
        """Execute transfer operation"""
        from_path = self.from_path.get()
        to_path = self.to_path.get()
        link_type = self.link_type.get()

        # Validate input
        if not from_path:
            messagebox.showerror("Error", self.t("error_from"))
            return

        if not to_path:
            messagebox.showerror("Error", self.t("error_to"))
            return

        if not os.path.exists(from_path):
            messagebox.showerror("Error", self.t("error_from_exist"))
            return

        # Get source name
        from_name = os.path.basename(from_path)
        to_full = os.path.join(to_path, from_name)

        # Check if target already exists
        if os.path.exists(to_full):
            messagebox.showerror("Error", self.t("error_to_exist") + f"\n{to_full}")
            return

        try:
            # First move source folder to target location
            shutil.move(from_path, to_full)

            # Create symbolic link (at original location, pointing to new location)
            # mklink command format: mklink [option] link_location target_location
            if link_type == "/D":
                cmd = f'mklink /D "{from_path}" "{to_full}"'
            else:  # /J
                cmd = f'mklink /J "{from_path}" "{to_full}"'

            # Execute mklink command (requires admin privileges)
            result = subprocess.run(cmd, capture_output=True, text=True, shell=True)

            if result.returncode == 0:
                # Add record
                record = {
                    "type": link_type,
                    "from": from_path,
                    "to": to_full,
                    "time": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                }
                self.records.append(record)
                self.save_records()
                self.refresh_records()

                messagebox.showinfo("Success", self.t("success").format(link_type, from_path, to_full))

                # Clear input
                self.from_path.set("")
                self.to_path.set("")
            else:
                # If link creation failed, try to restore file to original location
                try:
                    shutil.move(to_full, from_path)
                except:
                    pass
                messagebox.showerror("Error", self.t("error_link") + f"\n{result.stderr}")

        except Exception as e:
            messagebox.showerror("Error", str(e))

    def get_link_target(self, path):
        """Get the target of a symbolic link or junction"""
        try:
            # Use os.readlink for symbolic links
            if os.path.islink(path):
                return os.readlink(path)
            # For junctions, use fsutil command
            result = subprocess.run(['fsutil', 'reparsepoint', 'query', path],
                                   capture_output=True, text=True, creationflags=subprocess.CREATE_NO_WINDOW)
            if result.returncode == 0:
                # Parse the output to find the target path
                for line in result.stdout.split('\n'):
                    if 'Print Name:' in line:
                        # Extract the path from "Print Name: X:\path"
                        target = line.split(':', 1)[1].strip()
                        if target:
                            return target
            return None
        except:
            return None

    def add_record_dialog(self):
        """Add record dialog"""
        dialog = tk.Toplevel(self.root)
        dialog.title(self.t("dialog_title"))
        dialog.resizable(True, False)
        dialog.transient(self.root)
        dialog.grab_set()

        # Link type
        ttk.Label(dialog, text=self.t("dialog_link_type")).pack(anchor=tk.W, padx=20, pady=(20, 5))
        type_var = tk.StringVar(value="/J")
        type_frame = ttk.Frame(dialog)
        type_frame.pack(fill=tk.X, padx=20)
        ttk.Radiobutton(type_frame, text="/J", variable=type_var, value="/J").pack(side=tk.LEFT)
        ttk.Radiobutton(type_frame, text="/D", variable=type_var, value="/D").pack(side=tk.LEFT)

        # Source path (manual input only, no browse button)
        ttk.Label(dialog, text=self.t("dialog_from")).pack(anchor=tk.W, padx=20, pady=(10, 5))
        from_var = tk.StringVar()
        from_entry = ttk.Entry(dialog, textvariable=from_var, width=70)
        from_entry.pack(fill=tk.X, padx=20)

        # Target path (manual input only, no browse button)
        ttk.Label(dialog, text=self.t("dialog_to")).pack(anchor=tk.W, padx=20, pady=(10, 5))
        to_var = tk.StringVar()
        to_entry = ttk.Entry(dialog, textvariable=to_var, width=70)
        to_entry.pack(fill=tk.X, padx=20)

        # Button frame for Confirm and Cancel
        btn_frame = ttk.Frame(dialog)
        btn_frame.pack(pady=20)

        def confirm():
            if not from_var.get() or not to_var.get():
                messagebox.showerror("Error", self.t("dialog_error"))
                return

            record = {
                "type": type_var.get(),
                "from": from_var.get(),
                "to": to_var.get(),
                "time": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            }
            self.records.append(record)
            self.save_records()
            self.refresh_records()
            dialog.destroy()

        # Confirm button
        ttk.Button(btn_frame, text=self.t("dialog_confirm"), command=confirm).pack(side=tk.LEFT, padx=10)
        # Cancel button
        cancel_text = "Cancel" if self.lang.get() == "en" else "取消"
        ttk.Button(btn_frame, text=cancel_text, command=dialog.destroy).pack(side=tk.LEFT, padx=10)

        # Auto-size dialog after all widgets are packed
        dialog.update_idletasks()
        dialog.geometry(f"{dialog.winfo_reqwidth() + 40}x{dialog.winfo_reqheight() + 20}")

        # Center dialog relative to parent
        dialog.update_idletasks()
        pw = self.root.winfo_x() + self.root.winfo_width() // 2
        ph = self.root.winfo_y() + self.root.winfo_height() // 2
        dw = dialog.winfo_width()
        dh = dialog.winfo_height()
        dialog.geometry(f"+{pw - dw // 2}+{ph - dh // 2}")

    def delete_record(self):
        """Delete selected record"""
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("Warning", self.t("warning_select"))
            return

        if messagebox.askyesno("Confirm", self.t("confirm_delete")):
            # Collect record indices to delete (from item tags)
            indices_to_delete = set()
            for item in selected:
                tags = self.tree.item(item, "tags")
                for tag in tags:
                    if tag.startswith("rec_"):
                        indices_to_delete.add(int(tag[4:]))
            # Delete in reverse order
            for idx in sorted(indices_to_delete, reverse=True):
                if idx < len(self.records):
                    self.records.pop(idx)
            self.save_records()
            self.refresh_records()

    def clear_records(self):
        """Clear all records"""
        if messagebox.askyesno("Confirm", self.t("confirm_clear")):
            self.records.clear()
            self.save_records()
            self.refresh_records()

    def restore_single_record(self, record):
        """Restore a single record - returns (success, error_message)"""
        from_path = record.get("from", "")
        to_path = record.get("to", "")

        # Check if target exists
        if not os.path.exists(to_path):
            # Target doesn't exist, check source status
            from_is_link = os.path.islink(from_path)
            from_is_dir = os.path.isdir(from_path) and not from_is_link
            from_exists = os.path.exists(from_path)

            # If source also doesn't exist, the record is stale - mark as success to remove it
            if not from_exists:
                return True, None # Record is stale, just remove it

            # Source exists but target doesn't - can't restore
            return False, self.t("error_target_not_found") + f"\n{to_path}"

        # Check source status
        from_is_link = os.path.islink(from_path)
        from_is_dir = os.path.isdir(from_path) and not from_is_link
        from_exists = os.path.exists(from_path)

        # If source exists but is not a link/junction, we can't restore
        if from_exists and not from_is_link and not from_is_dir:
            # It's a file, not a directory/link
            return False, self.t("error_source_exists") + f"\n{from_path}"

        try:
            # Remove the symbolic link/junction at source if it exists
            if from_is_link:
                # For symbolic links, os.rmdir works
                os.rmdir(from_path)
            elif from_is_dir:
                # For junctions, need to use rmdir command
                subprocess.run(f'rmdir "{from_path}"', shell=True, check=True)
            # If source doesn't exist, that's fine - we'll just move the target

            # Move target folder back to source location
            shutil.move(to_path, from_path)
            return True, None

        except Exception as e:
            return False, str(e)

    def restore_selected(self):
        """Restore selected record(s)"""
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("Warning", self.t("warning_select_restore"))
            return
        
        # Get unique record indices from selection
        indices_to_restore = set()
        for item in selected:
            tags = self.tree.item(item, "tags")
            for tag in tags:
                if tag.startswith("rec_"):
                    indices_to_restore.add(int(tag[4:]))
        
        if not indices_to_restore:
            messagebox.showwarning("Warning", self.t("warning_select_restore"))
            return
        
        # If only one record selected, show detailed confirmation
        if len(indices_to_restore) == 1:
            idx = list(indices_to_restore)[0]
            if idx >= len(self.records):
                return
            record = self.records[idx]
            from_path = record.get("from", "")
            to_path = record.get("to", "")

            if not messagebox.askyesno("Confirm", self.t("confirm_restore_single").format(from_path, to_path)):
                return

            # Restore single record
            success, error = self.restore_single_record(record)

            if success:
                self.records.pop(idx)
                self.save_records()
                self.refresh_records()
                messagebox.showinfo("Success", self.t("restore_success_single").format(from_path, to_path))
            else:
                messagebox.showerror("Error", self.t("restore_failed").format(error))
        else:
            # Multiple records selected - confirm and restore all
            if not messagebox.askyesno("Confirm", self.t("confirm_restore")):
                return
            
            success_count = 0
            failed_count = 0
            errors = []
            
            # Sort indices in reverse to remove from list safely
            for idx in sorted(indices_to_restore, reverse=True):
                if idx < len(self.records):
                    record = self.records[idx]
                    success, error = self.restore_single_record(record)
                    
                    if success:
                        self.records.pop(idx)
                        success_count += 1
                    else:
                        failed_count += 1
                        errors.append(f"{record.get('from', '')}: {error}")
            
            self.save_records()
            self.refresh_records()
            
            if failed_count == 0:
                messagebox.showinfo("Success", self.t("restore_success").format(success_count))
            elif success_count == 0:
                messagebox.showerror("Error", self.t("restore_failed").format("\n".join(errors[:5])))
            else:
                messagebox.showinfo("Partial", self.t("restore_partial").format(success_count, failed_count))

    def restore_all(self):
        """Restore all links - remove symbolic links and move folders back"""
        if not self.records:
            messagebox.showinfo("Info", "No records to restore")
            return

        if not messagebox.askyesno("Confirm", self.t("confirm_restore")):
            return

        success_count = 0
        failed_count = 0
        errors = []

        for i in range(len(self.records) - 1, -1, -1):  # Iterate in reverse
            record = self.records[i]
            
            success, error = self.restore_single_record(record)
            
            if success:
                self.records.pop(i)
                success_count += 1
            else:
                failed_count += 1
                errors.append(f"{record.get('from', '')}: {error}")

        # Save updated records
        self.save_records()
        self.refresh_records()

        # Show result
        if failed_count == 0:
            messagebox.showinfo("Success", self.t("restore_success").format(success_count))
        elif success_count == 0:
            messagebox.showerror("Error", self.t("restore_failed").format("\n".join(errors[:5])))
        else:
            messagebox.showinfo("Partial", self.t("restore_partial").format(success_count, failed_count))

    def refresh_records(self):
        """Refresh records display"""
        # Clear existing items
        for item in self.tree.get_children():
            self.tree.delete(item)

        # Add records to tree: each record = 2 rows (source row + target row)
        for i, record in enumerate(self.records):
            link_type = record.get("type", "")
            from_path = record.get("from", "")
            to_path = record.get("to", "")
            rec_time = record.get("time", "")
            tag = f"rec_{i}"
            even_odd = "even" if i % 2 == 0 else "odd"

            # Row 1: name+type | source path
            # Get folder name from source path
            folder_name = os.path.basename(from_path) if from_path else ""
            self.tree.insert("", tk.END, values=(
                f"{folder_name} [{link_type}]",
                f"[F] {from_path}"
            ), tags=(tag, even_odd, "from_row"))

            # Row 2: empty | target path
            self.tree.insert("", tk.END, values=(
                "",
                f"[T] {to_path}"
            ), tags=(tag, even_odd, "to_row"))

        # Style alternating record groups
        self.tree.tag_configure("even", background="#f0f4ff")
        self.tree.tag_configure("odd", background="#ffffff")

    def load_records(self):
        """Load records"""
        try:
            if os.path.exists(self.records_file):
                with open(self.records_file, "r", encoding="utf-8") as f:
                    self.records = json.load(f)
        except Exception as e:
            print(f"Failed to load records: {e}")
            self.records = []

    def save_records(self):
        """Save records"""
        try:
            with open(self.records_file, "w", encoding="utf-8") as f:
                json.dump(self.records, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"Failed to save records: {e}")


def main():
    root = tk.Tk()
    app = MKLinkManager(root)
    root.mainloop()


if __name__ == "__main__":
    main()
