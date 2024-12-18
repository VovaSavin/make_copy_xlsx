import win32com.client
import shutil
import datetime

FILE_PATH = r"D:\ngu\test\Радіостанції_2024-11-23.xlsx"


def path_file_destination():
    fd = r"D:\ngu\test\Радіостанції_2024-11-23_copy.xlsx"
    name_file_destination = fd.split(f"\\")[-1].split(".")[0]
    new_name = f"{name_file_destination}_{str(datetime.datetime.today())}.xlsx".replace(
        " ", "_"
    ).replace(
        ":", "_"
    )
    return r"D:\ngu\test" + "\\" + new_name


FILE_DESTINATION = path_file_destination()


class SaveFile:
    def __init__(self, f_path: str, dispatch: win32com.client.Dispatch):
        self.f_path = f_path
        self.dispatch = dispatch
        self.active_objects = win32com.client.GetActiveObject("Excel.Application")

    def is_active(self, ):
        for x in self.active_objects.Workbooks:
            if x.FullName == self.f_path:
                return True
            return False

    def start_save(self):
        exl = self.dispatch
        if self.is_active():
            workbook_for = exl.Workbooks.Open(self.f_path)
            workbook_for.Save()
            workbook_for.Close()
        return None

    def run_file_after_saving(self):
        exl = self.dispatch
        if not self.is_active():
            exl.Workbooks.Open(self.f_path)
        return None

    def closer(self):
        exl = self.dispatch
        if self.is_active():
            workbook_for = exl.Workbooks.Open(self.f_path)
            workbook_for.Close()


class CopyFile:
    def __init__(self, sources: str, destination: str):
        self.sources = sources
        self.destination = destination

    def create_copy_file(self):
        shutil.copy2(self.sources, self.destination)


class BackSave(SaveFile, CopyFile):
    def __init__(self, sources, destination, dispatch, ):
        super().__init__(f_path=sources, dispatch=dispatch)
        super(SaveFile, self).__init__(sources=sources, destination=destination)


def run():
    bs = BackSave(FILE_PATH, FILE_DESTINATION, win32com.client.Dispatch("Excel.Application"), )

    bs.start_save()
    bs.create_copy_file()
    # bs.run_file_after_saving()
    bs.closer()


if __name__ == "__main__":
    run()
