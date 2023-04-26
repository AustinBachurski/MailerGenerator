import arcpy
from datetime import datetime


class Generator:
    def __init__(self, field_string, search_distance, search_string, map_name_string):
        self.fieldString = field_string.upper()
        self.gdb = "J:\\Austin\\Projects\\Mailing List\\Mailing List.gdb"
        self.field_values = []
        self.mailing_list_parcels = f"{self.gdb}\\MailingListParcels"
        self.map_name_string = map_name_string
        self.not_found = []
        self.other_parcels = "J:\\Austin\\County & Cadastral\\Clips\\MailingListReference.shp"
        self.project = arcpy.mp.ArcGISProject("J:\\Austin\\Projects\\Mailing List\\Mailing List.aprx")
        self.search_distance = f"{search_distance} Feet"
        self.search_string = search_string.upper()
        self.subject_parcel_path = f"{self.gdb}\\SubjectParcel"

    def append_mailing_list_parcels(self):
        arcpy.Append_management(inputs=self.mailing_lisp_parcels_selection(), target=self.mailing_list_parcels)

    def append_subject_parcel(self):
        arcpy.Append_management(inputs=self.subject_parcel_selection(), target=self.subject_parcel_path)

    def center_element(self, element):
        pageCenter = self.project.listLayouts("Mailing List Map")[0].pageWidth * .5
        return pageCenter - (element.elementWidth * .5)

    def clear_old_data(self):
        arcpy.DeleteFeatures_management(self.subject_parcel_path)
        arcpy.DeleteFeatures_management(self.mailing_list_parcels)

    def configure_map(self):
        arcpy.RecalculateFeatureClassExtent_management(self.mailing_list_parcels, store_extent=True)
        frame = self.project.listLayouts("Mailing List Map")[0].listElements("MAPFRAME_ELEMENT")[0]
        mailingListParcelsLayer = self.project.listMaps("Mailing List")[0].listLayers("Mailing List Parcels")[0]
        frame.camera.setExtent(frame.getLayerExtent(mailingListParcelsLayer,
                                                    selection_only=False, symbolized_extent=True))

        for element in self.project.listLayouts("Mailing List Map")[0].listElements("TEXT_ELEMENT"):
            if element.name == "Date":
                element.text = datetime.today().strftime("%B %#d, %Y")
            elif element.name == "MapName":
                element.text = self.map_name()
                element.elementPositionX = self.center_element(element)

    def export_pdf(self):
        arcpy.env.overwriteOutput = True
        self.project.listLayouts("Mailing List Map")[0].exportToPDF(
            f"J:\\Austin\\Maps\\Mailing List Map {self.map_name()}.pdf")

    def field_selection(self):
        select = {"ADDRESS": "AddressLin",
                  "TRACT ID": "TRACT_ID",
                  "ASSESSOR NUMBER": "ASSRNO"}

        return select[self.fieldString]

    def generate_spreadsheet(self):
        arcpy.env.overwriteOutput = True
        arcpy.TableToExcel_conversion(self.mailing_list_parcels,
                                      f"J:\\Austin\\Mailing Lists\\Mailing List {self.map_name()}.xlsx")

    def mailing_lisp_parcels_selection(self):
        return arcpy.SelectLayerByLocation_management(in_layer=self.other_parcels,
                                                      overlap_type="WITHIN_A_DISTANCE",
                                                      select_features=self.subject_parcel_path,
                                                      search_distance=self.search_distance)

    def map_name(self):
        if self.map_name_string:
            return self.map_name_string
        else:
            if ';' in self.search_string:
                inputs = self.search_string.split(';')
                return inputs[0]
            else:
                return ' '.join([word.capitalize() for word in self.search_string.split()]).lstrip("'").rstrip("'")

    def query_is_valid(self):
        self.not_found.clear()
        if not self.field_values:
            with arcpy.da.SearchCursor(self.other_parcels, self.field_selection()) as search_field:
                for value in search_field:
                    self.field_values.append((value[0]))

        for value in self.search_criteria():
            if value not in self.field_values:
                self.not_found.append(value)

        return not bool(self.not_found)

    def search_criteria(self):
        criteria = []
        if ';' in self.search_string:
            inputs = self.search_string.split(';')
            for each in inputs:
                criteria.append(each.upper())
        else:
            criteria.append(self.search_string.upper())
        return criteria

    def subject_parcel_selection(self):
        return arcpy.SelectLayerByAttribute_management(in_layer_or_view=self.other_parcels,
                                                       where_clause=self.sql_query())

    def sql_query(self):
        if ';' in self.search_string:
            expression = ""
            inputs = self.search_string.split(';')
            for each in inputs:
                expression += f"{self.field_selection()} = '{each.upper()}' or "
            return expression.rstrip(" or ")
        else:
            return f"{self.field_selection()} = '{self.search_string.upper()}'"


if __name__ == "__main__":

    field_input = input("Address, Tract ID, or Assessor Number?\n")
    distance = input("Include distance in feet?\n")
    search_criteria_input = input("Enter search criteria, separate multiple criteria with a ';'\n")
    map_name_input = input("Enter custom map name, if N/A leave blank.\n")

    step = Generator(field_input,
                     distance,
                     search_criteria_input,
                     map_name_input)
    if step.query_is_valid():
        step.clear_old_data()
        step.append_subject_parcel()
        step.append_mailing_list_parcels()
        step.generate_spreadsheet()
        step.configure_map()
        step.export_pdf()
    else:
        print(f"Search string(s) not found in field: {step.field_selection}")
        for not_found in step.not_found:
            print(not_found)
