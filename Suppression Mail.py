import imaplib
import tkinter as tk
from tkinter import messagebox
from datetime import datetime

class EmailDeleterApp:
    def __init__(self, master):
        self.master = master
        master.title("Suppression d'e-mails")
        master.geometry("400x400")

        self.label_username = tk.Label(master, text="Nom d'utilisateur:")
        self.label_password = tk.Label(master, text="Mot de passe:")
        self.label_start_date = tk.Label(master, text="Date de début (JJ/MM/AAAA):")
        self.label_end_date = tk.Label(master, text="Date de fin (JJ/MM/AAAA):")
        self.entry_username = tk.Entry(master)
        self.entry_password = tk.Entry(master, show="*")
        self.entry_start_date = tk.Entry(master)
        self.entry_end_date = tk.Entry(master)
        self.button_delete = tk.Button(master, text="Supprimer les e-mails entre les dates", command=self.delete_emails)
        self.button_stop = tk.Button(master, text="Arrêter", command=self.stop_deletion)

        self.label_username.grid(row=0, column=0, sticky="e")
        self.label_password.grid(row=1, column=0, sticky="e")
        self.label_start_date.grid(row=2, column=0, sticky="e")
        self.label_end_date.grid(row=3, column=0, sticky="e")
        self.entry_username.grid(row=0, column=1)
        self.entry_password.grid(row=1, column=1)
        self.entry_start_date.grid(row=2, column=1)
        self.entry_end_date.grid(row=3, column=1)
        self.button_delete.grid(row=4, column=0, columnspan=2, pady=10)
        self.button_stop.grid(row=5, column=0, columnspan=2)

        self.stopped = False

    def delete_emails(self):
        username = self.entry_username.get()
        password = self.entry_password.get()
        start_date_str = self.entry_start_date.get()
        end_date_str = self.entry_end_date.get()

        try:
            # Parsing des dates
            start_date = datetime.strptime(start_date_str, "%d/%m/%Y")
            end_date = datetime.strptime(end_date_str, "%d/%m/%Y")

            # Connexion SSL sécurisée
            mail = imaplib.IMAP4_SSL('outlook.office365.com', 993)

            # Authentification
            mail.login(username, password)

            # Sélectionner la boîte de réception
            mail.select('inbox')

            # Déterminer les dates limites
            start_date_limite = start_date.strftime('%d-%b-%Y')
            end_date_limite = end_date.strftime('%d-%b-%Y')

            # Recherche des e-mails entre les dates limites
            status, messages = mail.search(None, '(SINCE "%s" BEFORE "%s")' % (start_date_limite, end_date_limite))

            # Marquer les e-mails comme supprimés
            for num in messages[0].split():
                if self.stopped:
                    break
                mail.store(num, '+FLAGS', '\\Deleted')

            # Appliquer les suppressions
            mail.expunge()

            messagebox.showinfo("Information", "Les e-mails entre les dates spécifiées ont été supprimés.")
        except ValueError:
            messagebox.showerror("Erreur", "Format de date invalide. Utilisez JJ/MM/AAAA.")
        except imaplib.IMAP4.error as e:
            messagebox.showerror("Erreur", f"Erreur lors de la connexion: {str(e)}")
        finally:
            # Fermer la connexion
            mail.close()
            mail.logout()

    def stop_deletion(self):
        self.stopped = True
        messagebox.showinfo("Information", "Suppression d'e-mails arrêtée.")

root = tk.Tk()
app = EmailDeleterApp(root)
root.mainloop()
