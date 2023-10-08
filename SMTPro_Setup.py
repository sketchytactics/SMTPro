import customtkinter as ctk
import os
import pandas as pd
import yaml as ym

ctk.set_appearance_mode('dark')
ctk.set_default_color_theme('green')

setup_app = ctk.CTk()
setup_app.title('Setup Wizard')
setup_frame = ctk.CTkFrame(master=setup_app)
setup_frame.grid(padx=20, pady=20)
setup_title = ctk.CTkLabel(master=setup_frame, text='Setup', fg_color='transparent', text_color='white',
                           width=400, height=30, font=('Lucida Sans Unicode', 25))
setup_title.grid(row=0, column=0, columnspan=2, sticky='nsew')


def generate_config():
    smtp_server = smtp_server_entry.get()
    smtp_port = 587
    smtp_user = smtp_username_entry.get()
    smtp_password = smtp_password_entry.get()
    smtp_sender = smtp_sender_entry.get()
    if smtp_sender == '':
        smtp_sender = smtp_user
    config_data = {'smtp_server': smtp_server, 'smtp_port': smtp_port, 'smtp_user': smtp_user,
                   'smtp_password': smtp_password, 'smtp_sender': smtp_sender}
    config_directory = os.path.join(os.getcwd(), 'SMTProConfig')
    if not os.path.exists(config_directory):
        os.makedirs(config_directory)
    with open(os.path.join(config_directory, 'config.yaml'), 'w') as yaml_file:
        ym.dump(config_data, yaml_file, default_flow_style=False)
    outbox_template = pd.DataFrame(columns=['Invoice', 'Receiver Email', 'Receiver CC', 'Subject', 'Greeting',
                                            'Email Body 1', 'Email Body 2', 'Signature1', 'Signature2',
                                            'Attachment Name'])
    with pd.ExcelWriter('OutboxTemplateFile.xlsx', engine='xlsxwriter') as writer:
        outbox_template.to_excel(writer, index=False, sheet_name='Outbox')
        workbook = writer.book
        worksheet = writer.sheets['Outbox']
        header_format = workbook.add_format({'bold': True, 'align': 'center', 'border': 2, 'bg_color': '#D3D3D3'})
        for col_num, value in enumerate(outbox_template.columns.values):
            worksheet.write(0, col_num, value, header_format)
    outbox_attachment_path = os.path.join(os.getcwd(), 'OutboxAttachments')
    if not os.path.exists(outbox_attachment_path):
        os.makedirs(outbox_attachment_path)
    setup_app.destroy()


smtp_server_label = ctk.CTkLabel(master=setup_frame, text='Enter SMTP Server Address', fg_color='transparent',
                                 text_color='white', font=('Lucida Sans Unicode', 16), width=190, height=20)
smtp_server_entry = ctk.CTkEntry(master=setup_frame, placeholder_text='smtp.example.com', fg_color='transparent',
                                 text_color='white', font=('Lucida Sans Unicode', 14), width=190, height=20,
                                 placeholder_text_color='#D3D3D3')
smtp_server_label.grid(row=1, column=0, padx=5, pady=5)
smtp_server_entry.grid(row=1, column=1, padx=5, pady=5)
smtp_username_label = ctk.CTkLabel(master=setup_frame, text='Enter User Email Address', fg_color='transparent',
                                   text_color='white', font=('Lucida Sans Unicode', 16), width=190, height=20)
smtp_username_entry = ctk.CTkEntry(master=setup_frame, placeholder_text='email@domain.com', fg_color='transparent',
                                   text_color='white', font=('Lucida Sans Unicode', 14), width=190, height=20,
                                   placeholder_text_color='#D3D3D3')
smtp_username_label.grid(row=2, column=0, padx=5, pady=5)
smtp_username_entry.grid(row=2, column=1, padx=5, pady=5)
smtp_password_label = ctk.CTkLabel(master=setup_frame, text='Enter Email Password', fg_color='transparent',
                                   text_color='white', font=('Lucida Sans Unicode', 16), width=190, height=20)
smtp_password_entry = ctk.CTkEntry(master=setup_frame, placeholder_text='******', fg_color='transparent',
                                   text_color='white', font=('Lucida Sans Unicode', 14), width=190, height=20,
                                   placeholder_text_color='#D3D3D3')
smtp_password_label.grid(row=3, column=0, padx=5, pady=5)
smtp_password_entry.grid(row=3, column=1, padx=5, pady=5)
smtp_sender_label = ctk.CTkLabel(master=setup_frame, text='Enter sender email if different than user',
                                 fg_color='transparent', text_color='white', font=('Lucida Sans Unicode', 16),
                                 width=190, height=20)
smtp_sender_entry = ctk.CTkEntry(master=setup_frame, placeholder_text='email@domain.com', fg_color='transparent',
                                 text_color='white', font=('Lucida Sans Unicode', 14), width=190, height=20,
                                 placeholder_text_color='#D3D3D3')
smtp_sender_label.grid(row=4, column=0, padx=5, pady=5)
smtp_sender_entry.grid(row=4, column=1, padx=5, pady=5)
smtp_config_set = ctk.CTkButton(master=setup_frame, text='Set Config', width=80, height=20, fg_color='#0d6b18',
                                text_color='white', font=('Lucida Sans Unicode', 16), command=generate_config)
smtp_config_set.grid(row=5, column=0, columnspan=2, pady=10)

setup_app.mainloop()
