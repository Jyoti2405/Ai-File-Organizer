import os
import time
import shutil
from tkinter import *
from tkinter import ttk, filedialog, messagebox
from tkinter.scrolledtext import ScrolledText
from PIL import Image, ImageTk
import platform
import speech_recognition as sr

audio_formats = ['.mp3', '.wav']
video_formats = ['.mp4', '.mkv']
image_formats = ['.jpg', '.png']
doc_formats   = ['.pdf', '.docx', '.txt']
all_formats   = audio_formats + video_formats + image_formats + doc_formats
light_font = ('Arial', 12)
dark_font = ('Segoe UI', 12, 'bold')

file_icons = {}
UNDO_STACK = []

def main_page():
    global main, style, theme_var, icon_imgs
    main = Tk()
    main.title("File Organiser")
    main.geometry("1400x900")
    main.minsize(1100, 680)
    style = ttk.Style(main)
    style.theme_use('clam')
    main.configure(bg='#f6f6f6')
    theme_var = StringVar(value="light")
    icon_imgs = {}
    create_icons()
    create_gui()
    btn_toggle_theme = Button(main, text="🌙", command=toggle_theme, relief='flat', bg='#eaeaea')
    btn_toggle_theme.place(relx=0.96, rely=0.02, anchor='ne')
    main.mainloop()

def create_icons():
    for ext in ['folder', 'mp3', 'wav', 'mp4', 'mkv', 'jpg', 'png', 'pdf', 'docx', 'txt', 'file']:
        img = PhotoImage(width=16, height=16)
        c = {'mp3':'#42a5f5','wav':'#7e57c2','mp4':'#fbc02d','mkv':'#43a047','jpg':'#ef5350','png':'#66bb6a',
             'pdf':'#ff7043','docx':'#5c6bc0','txt':'#90a4ae','folder':'#ffd54f','file':'#bdbdbd'}[ext]
        img.put(c, to=(0,0,16,16))
        file_icons[ext] = img

def filetype_icon(path):
    ext = os.path.splitext(path)[1][1:].lower()
    return file_icons.get(ext, file_icons['file'])

def toggle_theme():
    if theme_var.get() == "light":
        style.theme_use('clam')
        style.configure('.', background='#212226', foreground='#f6f6f6',
                        fieldbackground='#232323', font=dark_font)
        main.configure(bg='#232323')
        update_widget_fonts(dark_font)
        theme_var.set("dark")
    else:
        style.theme_use('clam')
        style.configure('.', background='#f6f6f6', foreground='#232323',
                        fieldbackground='#fff', font=light_font)
        main.configure(bg='#f6f6f6')
        update_widget_fonts(light_font)
        theme_var.set("light")

def update_widget_fonts(new_font):
    def do_update(widget):
        try:
            widget.config(font=new_font)
        except Exception:
            pass
        for child in widget.winfo_children():
            do_update(child)
    do_update(main)

def create_gui():
    global contents, dest_str, params, org_feedback, selected_folders
    global organised_files_listbox, progress, sidebar, lbl_feedback, filter_var
    global frames, current_frame, preview_canvas, preview_text, preview_label, settings

    contents = {}
    selected_folders = []
    dest_str = StringVar()
    filter_var = StringVar()
    params = {}
    org_feedback = StringVar()
    settings = {'animation_speed':0.03, 'theme':'light'}
    frames = {}

    sidebar = Frame(main, width=120, bg='#dbeafe')
    sidebar.pack(side=LEFT, fill=Y)
    main_container = Frame(main, bg='#f6f6f6', width=1050)
    main_container.pack_propagate(False)
    main_container.pack(side=LEFT, fill=BOTH, expand=1, padx=(0,10))

    frames['source'] = Frame(main_container, bg='#f6f6f6')
    frames['destination'] = Frame(main_container, bg='#f6f6f6')
    frames['params'] = Frame(main_container, bg='#f6f6f6')
    frames['execute'] = Frame(main_container, bg='#f6f6f6')

    create_source_frame(frames['source'])
    create_destination_frame(frames['destination'])
    create_params_frame(frames['params'])
    create_execute_frame(frames['execute'])

    current_frame = frames['source']
    current_frame.pack(fill=BOTH, expand=1)

    add_sidebar_nav(sidebar)
    voice_btn = Button(sidebar, text="🎤 Voice Cmd", command=voice_command, bg="#f9c846", font=light_font)
    voice_btn.pack(fill=X, pady=(10,2), padx=8)

def show_frame(name):
    global current_frame
    current_frame.pack_forget()
    current_frame = frames[name]
    current_frame.pack(fill=BOTH, expand=1)

def add_sidebar_nav(frame):
    nav_items = [
        ("Source", lambda: show_frame('source'), 'folder'),
        ("Destination", lambda: show_frame('destination'), 'folder'),
        ("Params", lambda: show_frame('params'), 'folder'),
        ("Execute", lambda: show_frame('execute'), 'folder'),
        ("Settings", open_settings_dialog, 'folder'),
    ]
    for idx, (text, callback, icon) in enumerate(nav_items):
        btn = Button(frame, text=f" {text}", image=file_icons[icon], font=light_font, relief='flat', anchor='w',
                     compound=LEFT, padx=10, pady=10, bg="#dbeafe", bd=0, activebackground="#fee694", command=callback)
        btn.pack(fill=X, pady=(10 if idx==0 else 0, 2), padx=8)
        add_hover_effect(btn, '#fee694', "#dbeafe")

def add_hover_effect(widget, enter_color, leave_color):
    def on_enter(e): widget['bg'] = enter_color
    def on_leave(e): widget['bg'] = leave_color
    widget.bind('<Enter>', on_enter)
    widget.bind('<Leave>', on_leave)

def create_source_frame(frame):
    global selected_folders, source_tree
    Label(frame, text="Source folders", font=light_font, bg='#f6f6f6').grid(row=0, column=0, columnspan=3, padx=20, pady=18)
    frame.grid_rowconfigure(10, weight=1)
    source_tree = ttk.Treeview(frame, height=22)
    source_tree.heading("#0", text="Catalog", anchor=W)
    source_tree.column("#0", width=600)
    source_tree.grid(row=1, column=0, columnspan=2, padx=20, sticky="nsew")

    btn_browse = Button(frame, text="Add folder(s)", compound=LEFT, command=lambda: browse_folders(source_tree),
                        font=light_font, relief='groove', bg='#c7f0da', bd=0, activebackground='#a8d5ba')
    btn_browse.grid(row=2, column=0, padx=20, pady=24, sticky="w")
    add_hover_effect(btn_browse, '#a8d5ba', '#c7f0da')

    btn_remove = Button(frame, text="Remove", compound=LEFT, command=lambda: remove_selected(source_tree),
                        font=light_font, relief='groove', bg='#ffcdd2', bd=0, activebackground='#ff8888')
    btn_remove.grid(row=2, column=1, padx=20, pady=24, sticky="w")
    add_hover_effect(btn_remove, '#ff8888', '#ffcdd2')

    btn_next = Button(frame, text="Next", command=lambda: show_frame('destination'),
                      font=light_font, relief='groove', bg='#b8e0fe', bd=0, activebackground='#71bef2')
    btn_next.grid(row=3, column=2, padx=35, pady=24, sticky="e")
    add_hover_effect(btn_next, '#71bef2', '#b8e0fe')

def create_destination_frame(frame):
    Label(frame, text='Destination Folder', font=light_font, bg='#f6f6f6').grid(row=0, column=0, columnspan=3, pady=20)
    lbl_entry = Label(frame, text='Enter Destination:', font=light_font, bg='#f6f6f6')
    lbl_entry.grid(row=1, column=0, pady=10, padx=10, sticky="e")
    entry_dest = Entry(frame, textvariable=dest_str, font=light_font, width=65, state="normal")
    entry_dest.grid(row=1, column=1, pady=10, padx=10, sticky="w")
    btn_browse_dest = Button(frame, text="Browse", compound=LEFT, command=select_dest,
                            font=light_font, relief='groove', bg='#c7f0da', bd=0, activebackground='#a8d5ba')
    btn_browse_dest.grid(row=1, column=2, padx=10, pady=10, sticky="w")
    add_hover_effect(btn_browse_dest, '#a8d5ba', '#c7f0da')

    btn_next = Button(frame, text="Next", command=lambda: show_frame('params'),
                      font=light_font, relief='groove', bg='#b8e0fe', bd=0, activebackground='#71bef2')
    btn_next.grid(row=3, column=2, padx=35, pady=24, sticky="e")
    add_hover_effect(btn_next, '#71bef2', '#b8e0fe')

def create_params_frame(frame):
    Label(frame, text="Select file types to organize:", font=light_font, bg='#f6f6f6').grid(row=0, column=0, padx=10, pady=12, columnspan=2)
    params['audio'] = IntVar(value=1)
    params['video'] = IntVar(value=1)
    params['image'] = IntVar(value=1)
    params['doc']   = IntVar(value=1)
    Checkbutton(frame, text="Audio Files (.mp3, .wav)", variable=params['audio'], font=light_font, bg="#f6f6f6").grid(row=1, column=0, sticky='w', padx=30)
    Checkbutton(frame, text="Video Files (.mp4, .mkv)", variable=params['video'], font=light_font, bg="#f6f6f6").grid(row=2, column=0, sticky='w', padx=30)
    Checkbutton(frame, text="Image Files (.jpg, .png)", variable=params['image'], font=light_font, bg="#f6f6f6").grid(row=3, column=0, sticky='w', padx=30)
    Checkbutton(frame, text="Documents (.pdf, .docx, .txt)", variable=params['doc'], font=light_font, bg="#f6f6f6").grid(row=4, column=0, sticky='w', padx=30)

    btn_next = Button(frame, text="Next", command=lambda: show_frame('execute'),
                      font=light_font, relief='groove', bg='#b8e0fe', bd=0, activebackground='#71bef2')
    btn_next.grid(row=5, column=1, padx=35, pady=24, sticky="e")
    add_hover_effect(btn_next, '#71bef2', '#b8e0fe')

def create_execute_frame(frame):
    global organised_files_listbox, progress, lbl_feedback, preview_canvas, preview_text, preview_label, filter_var
    lbl_exec = Label(frame, text="Click below to organize files as per selected parameters.", font=light_font, bg='#f6f6f6')
    lbl_exec.pack(pady=10)
    top_row = Frame(frame, bg='#f6f6f6')
    top_row.pack(fill=X)
    filter_entry = Entry(top_row, textvariable=filter_var)
    filter_entry.pack(side=RIGHT, padx=10)
    filter_btn = Button(top_row, text="Filter", command=filter_list)
    filter_btn.pack(side=RIGHT)
    btn_organize = Button(frame, text="Organize Files", compound=LEFT, command=organize_files,
                          font=light_font, relief='groove', bg='#77dd77', bd=0, activebackground='#38a169')
    btn_organize.pack(pady=4)
    add_hover_effect(btn_organize, '#38a169', '#77dd77')
    undo_btn = Button(frame, text="Undo", command=undo_last_move, font=light_font)
    undo_btn.pack()
    lbl_feedback = Label(frame, textvariable=org_feedback, font=light_font, fg='green', bg='#f6f6f6')
    lbl_feedback.pack()
    progress = ttk.Progressbar(frame, orient=HORIZONTAL, length=700, mode='determinate')
    progress.pack(pady=4)
    progress['value'] = 0
    bottom_frame = Frame(frame)
    bottom_frame.pack(fill=BOTH, expand=1)
    scrollbar_filelist = Scrollbar(bottom_frame)
    organised_files_listbox = Listbox(bottom_frame, yscrollcommand=scrollbar_filelist.set, width=65, height=18, font=light_font)
    organised_files_listbox.pack(side=LEFT, fill=Y)
    scrollbar_filelist.pack(side=LEFT, fill=Y)
    scrollbar_filelist.config(command=organised_files_listbox.yview)
    organised_files_listbox.bind('<Button-3>', right_click_menu)
    organised_files_listbox.bind('<<ListboxSelect>>', show_file_preview)
    prevw_frame = Frame(bottom_frame)
    prevw_frame.pack(side=LEFT, fill=BOTH, expand=1)
    preview_label = Label(prevw_frame, text="Preview", font=dark_font)
    preview_label.pack(pady=(10,0))
    preview_canvas = Canvas(prevw_frame, width=280, height=180, bg='#eee')
    preview_canvas.pack()
    preview_text = ScrolledText(prevw_frame, width=36, height=7, font=('Consolas', 11))
    preview_text.pack()

def voice_command():
    org_feedback.set("Listening for command...")
    main.update_idletasks()
    recognizer = sr.Recognizer()
    with sr.Microphone() as source:
        try:
            audio = recognizer.listen(source, timeout=4)
            command = recognizer.recognize_google(audio)
            org_feedback.set(f"Recognized: {command}")
            handle_voice_action(command.lower().strip())
        except sr.WaitTimeoutError:
            org_feedback.set("Mic timeout, try again.")
        except sr.UnknownValueError:
            org_feedback.set("Could not understand.")
        except Exception as e:
            org_feedback.set(str(e))

def handle_voice_action(cmd):
    if "undo" in cmd:
        undo_last_move()
    elif cmd.startswith("open file "):
        filename = cmd.replace("open file ", "").strip()
        open_file_by_voice(filename)
    elif cmd.startswith("open folder "):
        foldername = cmd.replace("open folder ", "").strip()
        open_folder_by_voice(foldername)
    elif "source" in cmd:
        show_frame('source')
    elif "destination" in cmd:
        show_frame('destination')
    elif "parameter" in cmd or "setting" in cmd:
        show_frame('params')
    elif "execute" in cmd or "result" in cmd:
        show_frame('execute')
    elif "settings" in cmd:
        open_settings_dialog()
    else:
        org_feedback.set("No matching command.")

def open_file_by_voice(filename):
    found = None
    for i in range(organised_files_listbox.size()):
        path = organised_files_listbox.get(i)
        if os.path.basename(path).lower() == filename.lower():
            found = path
            break
    if found:
        open_file(found)
        org_feedback.set(f"Opened file: {filename}")
    else:
        dest_root = dest_str.get()
        for root, dirs, files in os.walk(dest_root):
            for f in files:
                if f.lower() == filename.lower():
                    open_file(os.path.join(root, f))
                    org_feedback.set(f"Opened file: {filename}")
                    return
        org_feedback.set("File not found.")

def open_folder_by_voice(foldername):
    dest_root = dest_str.get()
    found = None
    for root, dirs, files in os.walk(dest_root):
        for d in dirs:
            if d.lower() == foldername.lower():
                found = os.path.join(root, d)
                break
    if found:
        open_folder(found)
        org_feedback.set(f"Opened folder: {foldername}")
    else:
        org_feedback.set("Folder not found.")

def browse_folders(tree):
    folder = filedialog.askdirectory()
    if folder and folder not in selected_folders:
        tree.insert('', 'end', text=folder)
        selected_folders.append(folder)

def remove_selected(tree):
    for item in tree.selection():
        folder = tree.item(item, "text")
        if folder in selected_folders:
            selected_folders.remove(folder)
        tree.delete(item)

def select_dest():
    folder = filedialog.askdirectory()
    if folder:
        dest_str.set(folder)

def count_files():
    selected_types = {
        'audio': params['audio'].get(),
        'video': params['video'].get(),
        'image': params['image'].get(),
        'doc':   params['doc'].get()
    }
    type_map = [
        ('audio', audio_formats),
        ('video', video_formats),
        ('image', image_formats),
        ('doc',   doc_formats)
    ]
    file_count = 0
    for folder in selected_folders:
        for root, dirs, files in os.walk(folder):
            for f in files:
                for k, exts in type_map:
                    if selected_types[k] and any(f.lower().endswith(ext) for ext in exts):
                        file_count += 1
                        break
    return file_count

def animate_feedback(color="#38a169"):
    for i in range(2):
        lbl_feedback.config(fg=color)
        main.update_idletasks()
        time.sleep(0.18)
        lbl_feedback.config(fg="#eee")
        main.update_idletasks()
        time.sleep(0.18)
    lbl_feedback.config(fg=color)
    main.update_idletasks()

def organize_files():
    progress['value'] = 0
    dest = dest_str.get()
    if not dest or not selected_folders:
        org_feedback.set("Please select both source and destination folders!")
        animate_feedback("#e53e3e")
        return
    organised_files_listbox.delete(0, END)
    selected_types = {
        'audio': params['audio'].get(),
        'video': params['video'].get(),
        'image': params['image'].get(),
        'doc':   params['doc'].get()
    }
    type_map = [
        ('audio', audio_formats, "Audio"),
        ('video', video_formats, "Video"),
        ('image', image_formats, "Image"),
        ('doc',   doc_formats,   "Documents")
    ]
    moved_files = []
    try:
        total_files = count_files()
        progress['maximum'] = total_files if total_files else 1
        count = 0
        for folder in selected_folders:
            for root, dirs, files in os.walk(folder):
                for f in files:
                    fpath = os.path.join(root, f)
                    for k, exts, foldername in type_map:
                        if selected_types[k] and any(f.lower().endswith(ext) for ext in exts):
                            extension = os.path.splitext(f)[1][1:]
                            dest_folder = os.path.join(dest, foldername, extension)
                            os.makedirs(dest_folder, exist_ok=True)
                            new_path = os.path.join(dest_folder, f)
                            if os.path.exists(new_path):   # Duplicate handler
                                new_path = handle_duplicate(new_path)
                            shutil.move(fpath, new_path)
                            organised_files_listbox.insert(END, new_path)
                            moved_files.append((new_path, fpath))
                            count += 1
                            progress['value'] = count
                            main.update_idletasks()
                            time.sleep(float(settings['animation_speed']))
                            break
        if count > 0:
            org_feedback.set(f"{count} files organized successfully!")
            UNDO_STACK.append(moved_files)
            animate_feedback("#38a169")
        else:
            org_feedback.set("No files matched selected parameters.")
            animate_feedback("#e53e3e")
        progress['value'] = 0
    except Exception as e:
        org_feedback.set(f"Error: {e}")
        animate_feedback("#e53e3e")
        progress['value'] = 0

def handle_duplicate(fname):
    name, ext = os.path.splitext(fname)
    j = 1
    while os.path.exists(f"{name}_copy{j}{ext}"):
        j += 1
    return f"{name}_copy{j}{ext}"

def undo_last_move():
    if not UNDO_STACK: return
    last = UNDO_STACK.pop()
    for to, frm in reversed(last):
        try:
            shutil.move(to, frm)
        except Exception: pass
    org_feedback.set("Undo complete")
    filter_list()

def right_click_menu(event):
    lbwidget = event.widget
    try:
        idx = lbwidget.nearest(event.y)
        lbwidget.select_clear(0, END)
        lbwidget.select_set(idx)
        lbwidget.activate(idx)
        file_path = lbwidget.get(idx)
    except:
        return
    menu = Menu(main, tearoff=0)
    menu.add_command(label="Open File", command=lambda: open_file(file_path))
    menu.add_command(label="Open Folder", command=lambda: open_folder(file_path))
    menu.add_command(label="Copy Path", command=lambda: main.clipboard_append(file_path))
    menu.tk_popup(event.x_root, event.y_root)

def open_file(fp):
    try:
        if platform.system() == 'Windows':
            os.startfile(fp)
        elif platform.system() == 'Darwin':
            os.system(f"open '{fp}'")
        else:
            os.system(f"xdg-open '{fp}'")
    except Exception as e:
        messagebox.showerror("Open Error", str(e))

def open_folder(fp):
    folder = os.path.dirname(fp)
    open_file(folder)

def show_file_preview(event):
    sel = event.widget.curselection()
    if not sel: return
    path = event.widget.get(sel[0])
    ext = os.path.splitext(path)[1].lower()
    preview_canvas.delete("all")
    preview_text.delete('1.0', END)
    preview_label.config(text=os.path.basename(path))
    if ext in ['.jpg', '.png']:
        try:
            img = Image.open(path)
            img.thumbnail((260,180))
            img = ImageTk.PhotoImage(img)
            preview_canvas.img = img
            preview_canvas.create_image(140,90, image=img)
        except Exception as e:
            preview_text.insert(END, f"Image Error: {e}")
    elif ext in ['.txt']:
        try:
            with open(path,"r",encoding='utf-8',errors='ignore') as f:
                txt = f.read(1000)
            preview_text.insert(END, txt)
        except Exception as e:
            preview_text.insert(END, f"Text Error: {e}")
    else:
        preview_text.insert(END,"No Preview Available")

def filter_list():
    val = filter_var.get().strip()
    organised_files_listbox.delete(0, END)
    for root, dirs, files in os.walk(dest_str.get()):
        for f in files:
            if (not val) or (val.lower() in f.lower()):
                path = os.path.join(root, f)
                if os.path.splitext(f)[1].lower() in all_formats:
                    organised_files_listbox.insert(END, path)

def open_settings_dialog():
    def save_settings():
        try:
            settings['animation_speed'] = float(speed_var.get())
        except Exception: pass
        win.destroy()
        org_feedback.set("Settings saved")
    win = Toplevel(main)
    win.title("Settings")
    Label(win,text="Animation Speed (seconds):").pack()
    speed_var=StringVar(value=str(settings['animation_speed']))
    Entry(win,textvariable=speed_var).pack()
    Button(win,text='Save',command=save_settings).pack()

if __name__ == "__main__":
    main_page()
