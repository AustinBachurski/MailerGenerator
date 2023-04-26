import customtkinter as ctk
import win11toast
from showinfm import show_in_file_manager


class Mailer:
    def __init__(self, window):
        # Query Frame
        self.query_frame = ctk.CTkFrame(master=window)
        self.query_frame.pack(fill=ctk.X, padx=5, pady=5)
        self.query_instructions = ctk.CTkLabel(master=self.query_frame,
                                               text="If you have multiple search criteria, "
                                               "separate them with a semicolon as displayed below.").pack()
        self.field_option = ctk.CTkOptionMenu(master=self.query_frame,
                                              values=["Address", "Tract ID", "Assessor Number"])
        self.field_option.pack(side=ctk.LEFT)
        self.distance_option = ctk.CTkOptionMenu(master=self.query_frame, values=["150 Feet", "225 Feet"])
        self.distance_option.pack(side=ctk.LEFT, padx=5)
        self.query_box = ctk.CTkEntry(master=self.query_frame,
                                      placeholder_text="Example Address 1;Example Address 2")
        self.query_box.pack(side=ctk.LEFT, fill=ctk.X, expand=True)
        # Map Frame
        self.name_frame = ctk.CTkFrame(master=window)
        self.name_frame.pack(fill=ctk.X, padx=5, pady=5)
        self.custom_name = ctk.CTkEntry(master=self.name_frame, state=ctk.DISABLED)
        self.checkbox_state = ctk.BooleanVar()
        self.use_custom = ctk.CTkCheckBox(master=self.name_frame,
                                          text="Use Custom Map Name ->",
                                          command=self.activate_custom_name,
                                          variable=self.checkbox_state,
                                          onvalue=True,
                                          offvalue=False)
        self.use_custom.pack(side=ctk.LEFT)
        self.custom_name.pack(side=ctk.LEFT, padx=5, fill=ctk.X, expand=True)
        self.generate_button = ctk.CTkButton(master=self.name_frame,
                                             text="Generate Mailer",
                                             command=self.display_result)
        self.generate_button.pack(side=ctk.RIGHT)

    def activate_custom_name(self):
        if self.checkbox_state.get():
            self.custom_name.configure(state=ctk.NORMAL)
        elif not self.checkbox_state.get():
            self.custom_name.configure(state=ctk.DISABLED)

    def display_result(self):
        if not self.query_box.get():
            self.query_box.configure(placeholder_text="You must enter your search criteria!")
        else:
            result, name = self.run_generator()
            error = ";".join(result)
            if error:
                win11toast.toast(f"Mailer Error: Search string was not found!\n"
                                 f"Field was: {self.field_option.get()}\n"
                                 f"Search string was: {error}")
            else:
                win11toast.toast(f"Mailer has been generated as:\n{name}",
                                 on_click=lambda args: self.show_files(args, name))

    @staticmethod
    def show_files(_, name):
        show_in_file_manager(f"J:\\Austin\\Maps\\Mailing List Map {name}.pdf")
        show_in_file_manager(f"J:\\Austin\\Mailing Lists\\Mailing List {name}.xlsx")

    def run_generator(self):
        import MailerGenerator
        generator = MailerGenerator.Generator(self.field_option.get(),
                                              self.distance_option.get(),
                                              self.query_box.get(),
                                              self.custom_name.get())
        if generator.query_is_valid():
            generator.clear_old_data()
            generator.append_subject_parcel()
            generator.append_mailing_list_parcels()
            generator.generate_spreadsheet()
            generator.configure_map()
            generator.export_pdf()
            # not_found is a list that is empty by default.
            return generator.not_found, generator.map_name()
        else:
            return generator.not_found, generator.map_name()


def main():
    root = ctk.CTk()
    root.title("Mailer Generator")
    root.geometry("600x110")
    Mailer(root)
    root.mainloop()


if __name__ == "__main__":
    main()
