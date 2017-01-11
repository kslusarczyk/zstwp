import xlrd
import csv


class Alarms_Correlation:

    def __init__(self):
        self.xls_file = 'HuaweiU2000Export_sorted_by_event_time.xlsx'
        self.csv_name = 'HuaweiU2000Export_sorted_by_event_time.csv'
        self.sheet_name = 'AlarmExport'
        self.notification_identifier = []
        self.acknowledge = []
        self.acknowledge_time = []
        self.acknowledge_user_id = []
        self.adapter_name = []
        self.adapter_specific_info = []
        self.additional_information = []
        self.additional_information_two = []
        self.additional_text = []
        self.alarm_group_name = []
        self.alarm_update_count = []
        self.base_time = []
        self.cleared = []
        self.clear_time = []
        self.create_date = []
        self.displayed_text = []
        self.event_action_log = []
        self.event_classification = []
        self.event_counter = []
        self.event_qualification = []
        self.event_time = []
        self.event_type = []
        self.internal_category = []
        self.invoke_identifier = []
        self.ip_time = []
        self.managed_object_class = []
        self.mo_identifier = []
        self.modify_date = []
        self.nem_notification_ident = []
        self.note = []
        self.notes_present = []
        self.original_severity = []
        self.perceived_severity = []
        self.potential_root_cause_flag = []
        self.priority = []
        self.priority_level = []
        self.probable_cause = []
        self.probable_cause_id = []
        self.proposed_repair_actions = []
        self.qualified_time = []
        self.reference_id = []
        self.sae_flag = []
        self.specific_problems = []
        self.terminate_time = []
        self.threshold_information = []
        self.trend_indication = []
        self.unknown_event_flag = []
        self.event_received_time = []
        self.mo_component = []
        self.next_change_date = []
        self.first_creation_time = []
        self.tagged_data = []


    # "NOTIFICATION_IDENTIFIER", "ACKNOWLEDGE", "ACKNOWLEDGE_TIME", "ACKNOWLEDGE_USER_ID", "ADAPTER_NAME",
    # "ADAPTER_SPECIFIC_INFO", "ADDITIONAL_INFORMATION", "ADDITIONAL_INFORMATION_TWO", "ADDITIONAL_TEXT",
    # "ALARM_GROUP_NAME", "ALARM_UPDATE_COUNT", "BASE_TIME", "CLEARED", "CLEAR_TIME", "EVENT_COUNTER",
    # "EVENT_QUALIFICATION","EVENT_TIME", "EVENT_TYPE", "INTERNAL_CATEGORY", "INVOKE_IDENTIFIER", "IP_TIME",
    # "MANAGED_OBJECT_CLASS", "MO_IDENTIFIER","MODIFY_DATE","NEM_NOTIFICATION_IDENT", "NOTE", "NOTES_PRESENT",
    # "ORIGINAL_SEVERITY", "PERCEIVED_SEVERITY", "POTENTIAL_ROOT_CAUSE_FLAG","PRIORITY", "PRIORITY_LEVEL",
    # "PROBABLE_CAUSE", "PROBABLE_CAUSE_ID","PROPOSED_REPAIR_ACTIONS", "QUALIFIED_TIME","REFERENCE_ID", "SAE_FLAG",
    # "SPECIFIC_PROBLEMS", "TERMINATE_TIME","THRESHOLD_INFORMATION", "TREND_INDICATION", "UNKNOWN_EVENT_FLAG",
    # "EVENT_RECEIVED_TIME", "MO_COMPONENT", "NEXT_CHANGE_DATE", "FIRST_CREATION_TIME"

    def csv_from_excel(self):
        b = xlrd.open_workbook(self.xls_file)
        s = b.sheet_by_name(self.sheet_name)
        bc = open(self.csv_name, 'w')
        bcw = csv.writer(bc, csv.excel)
        for row in range(s.nrows):
            this_row = []
            for col in range(s.ncols):
                val = s.cell_value(row, col)
                if isinstance(val, unicode):
                    val = val.encode('utf8')
                this_row.append(val)
            bcw.writerow(this_row)

    def load_data_categories(self):
        with open(self.csv_name, 'r') as csv_file:
            self.csv_file = csv.reader(csv_file, delimiter = ',')
            for row in self.csv_file:
                self.notification_identifier.append(row[0]), self.acknowledge.append(row[1]), self.acknowledge_time.append(row[2]),
                self.acknowledge_user_id.append(row[3]), self.adapter_name.append(row[4]), self.adapter_specific_info.append(row[5]),
                self.additional_information.append(row[6]), self.additional_information_two.append(row[7]),
                self.additional_text.append(row[8]), self.alarm_group_name.append(row[9]), self.alarm_update_count.append(row[10]),
                self.base_time.append(row[11]), self.cleared.append(row[12]), self.clear_time.append(row[13]),
                self.create_date.append(row[14]), self.displayed_text.append(row[15]), self.event_action_log.append(row[16]),
                self.event_classification.append(row[17]), self.event_counter.append(row[18]), self.event_qualification.append(row[19]),
                self.event_time.append(row[20]), self.event_type.append(row[21]), self.internal_category.append(row[22]),
                self.invoke_identifier.append(row[23]), self.ip_time.append(row[24]), self.managed_object_class.append(row[25]),
                self.mo_identifier.append(row[26]), self.modify_date.append(row[27]), self.note.append(row[29]),
                self.notes_present.append(row[30]), self.original_severity.append(row[31]), self.perceived_severity.append(row[32]),
                self.potential_root_cause_flag.append(row[33]), self.priority.append(row[34]), self.priority_level.append(row[35]),
                self.probable_cause.append(row[36]), self.probable_cause_id.append([37]), self.proposed_repair_actions.append(row[38]),
                self.qualified_time.append(row[39]), self.reference_id.append(row[40]), self.sae_flag.append(row[41]),
                self.specific_problems.append(row[42]), self.terminate_time.append(row[43]), self.threshold_information.append(row[44]),
                self.trend_indication.append(row[45]), self.unknown_event_flag.append(row[46]), self.event_received_time.append(row[47]),
                self.mo_component.append(row[48]), self.next_change_date.append(row[49]), self.first_creation_time.append(row[50])


alarms_correlation = Alarms_Correlation()
#alarms_correlation.csv_from_excel()
alarms_correlation.load_data_categories()
