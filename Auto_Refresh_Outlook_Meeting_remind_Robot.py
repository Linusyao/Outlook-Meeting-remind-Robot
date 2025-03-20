import win32com.client
import pythoncom
import datetime
import tkinter as tk
import threading

class OutlookMeetingReminder:
    def __init__(self, refresh_interval=3600):
        self.refresh_interval = refresh_interval
        self.root = tk.Tk()
        self.root.title("Heute Outlook Meeting")
        self.root.geometry("300x200+1200+700")
        self.root.attributes('-topmost', True)
        self.root.resizable(False, False)
        self.frame = tk.Frame(self.root)
        self.frame.pack(pady=5)
        tk.Label(self.frame, text="Heute Outlook Meeting Reminder", font=("Arial", 12, "bold")).pack()
        self.meeting_labels = []
        self.update_meetings()

    def get_today_outlook_calendar(self):
        pythoncom.CoInitialize()
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        calendar = outlook.GetDefaultFolder(9)
        today = datetime.datetime.now()
        begin = today.replace(hour=0, minute=0, second=0)
        end = today.replace(hour=23, minute=59, second=59)
        restriction = f"[Start] >= '{begin.strftime('%m/%d/%Y %H:%M %p')}' AND [Start] <= '{end.strftime('%m/%d/%Y %H:%M %p')}'"
        items = calendar.Items
        items.IncludeRecurrences = True
        items.Sort("[Start]")
        today_items = items.Restrict(restriction)
        meetings = []
        for appointment in today_items:
            subject = appointment.Subject
            start = appointment.Start.strftime('%H:%M')
            end_time = appointment.End.strftime('%H:%M')
            meetings.append(f"{start}-{end_time} {subject}")
        return meetings

    def update_meetings(self):
        for label in self.meeting_labels:
            label.destroy()
        self.meeting_labels.clear()
        meetings = self.get_today_outlook_calendar()
        if not meetings:
            lbl = tk.Label(self.frame, text="Keine Meetings heute", font=("Arial", 10))
            lbl.pack()
            self.meeting_labels.append(lbl)
        else:
            for m in meetings:
                lbl = tk.Label(self.frame, text=m, font=("Arial", 10), anchor="w", justify="left")
                lbl.pack()
                self.meeting_labels.append(lbl)
        self.root.after(self.refresh_interval * 1000, self.update_meetings)

    def run(self):
        self.root.mainloop()

def main():
    app = OutlookMeetingReminder(refresh_interval=3600)
    app.run()

threading.Thread(target=main).start()
