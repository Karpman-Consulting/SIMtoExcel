import re
import os
import tkinter as tk
from tkinter import filedialog
import xlsxwriter
import math

import add_logo


class SIMFileReader:
    def __init__(self, sim_file_path):
        self.file_path = sim_file_path
        self.file_name = "".join(os.path.basename(self.file_path).split('.')[:-1])
        self.wb_name = os.path.join(os.path.dirname(self.file_path), f'{self.file_name} - SIM.xlsx')
        self.report_contents = {}
        self.doe_version = None

        self.bepu_data = None
        self.ls_b_data = None
        self.lv_b_data = None
        self.lv_d_data = None
        self.ps_c_data = None
        self.pv_a_data = None
        self.ss_a_data = None
        self.ss_b_data = None
        self.ss_f_data = None
        self.ss_g_data = None
        self.ss_h_data = None
        self.ss_l_data = None
        self.ss_r_data = None
        self.sv_a_data = None

        self.parsing_methods = {
            "BEPU": "parse_bepu",
            "LS-B": "parse_ls_b",
            "LV-B": "parse_lv_b",
            "LV-D": "parse_lv_d",
            "PS-C": "parse_ps_c",
            "PV-A": "parse_pv_a",
            "SS-A": "parse_ss_a",
            "SS-B": "parse_ss_b",
            "SS-F": "parse_ss_f",
            "SS-G": "parse_ss_g",
            "SS-H": "parse_ss_h",
            "SS-L": "parse_ss_l",
            "SS-R": "parse_ss_r",
            "SV-A": "parse_sv_a",
        }

    @staticmethod
    def clean(val):
        try:
            return float(val)
        except ValueError:
            return ""

    def read_file(self):
        with open(self.file_path, "r", encoding="iso-8859-1") as f:
            sim_contents = f.read().splitlines()

        active_report = None
        system_zone_or_space = ""
        active_report_contents = []
        i = 0
        while i < len(sim_contents):
            line = sim_contents[i]
            if "BDL RUN" in line:
                if not self.doe_version:
                    self.doe_version = line[81:88]

                i += 2  # skip next line
                continue

            if "REPORT- " in line:
                index = line.index("REPORT- ") + len("REPORT- ")
                report = line[index:index + 4]

                # These reports occur per system/zone/space, so we need to parse the system/zone/space name
                if report in ["LS-B", "SS-A", "SS-B", "SS-F", "SS-G", "SS-H", "SS-L", "SS-R", "SV-A"]:
                    parts: list[str] = re.split(r"\s{2,}", line)
                    system_zone_or_space = parts[1].strip()
                if active_report is None:
                    active_report = report
                if report != active_report:
                    if active_report not in self.report_contents:
                        self.report_contents[active_report] = active_report_contents
                    else:
                        self.report_contents[active_report].extend(active_report_contents)

                    active_report_contents = []
                    active_report = report

            elif active_report is not None:
                # These reports need the system/zone parsed from the "REPORT- " line
                if active_report in ["LS-B", "SS-A", "SS-B", "SS-F", "SS-G", "SS-H", "SS-L", "SS-R", "SV-A"]:
                    active_report_contents.append((system_zone_or_space, line))
                else:
                    active_report_contents.append(line)

            i += 1

    def parse_contents(self):
        for report in self.report_contents:
            if report in self.parsing_methods:
                report_parse_method = getattr(self, self.parsing_methods[report])
                report_parse_method()

    def parse_bepu(self):
        skipline_substrings = [
            "------",
            "BDL RUN",
        ]
        end_use_spans = [(4, 12), (12, 21), (21, 30), (30, 39), (39, 48), (48, 57), (57, 66), (66, 75), (75, 84),
                         (84, 93), (93, 102), (102, 111), (111, 120), (120, 130)]

        lines = self.report_contents['BEPU']

        summary_section = False
        data = [
            ["Meter", "Units", "Lights", "Task Lights", "Misc. Equip.", "Heating", "Cooling", "Heat Rejection", "Pumps & Aux.", "Fans", "Rerig. Display", "Ht. Pump Supplemental", "Domestic Hot Water", "Exterior", "Total"]
        ]
        row = []
        meter_data_lines_ctr = 0
        for line in lines[6:]:

            if len(line.strip()) == 0 or any(substring in line for substring in skipline_substrings):
                continue

            if "TOTAL ELECTRICITY" in line:
                summary_section = True

            if not summary_section:
                if meter_data_lines_ctr % 2 == 0:
                    row.append(line.strip())
                else:
                    # Parse and clean the data using the spans then append to row
                    for i, (start, end) in enumerate(end_use_spans):
                        if i == 0:
                            row.append(line[start:end].strip())
                        else:
                            row.append(self.clean(line[start:end].strip()))
                    data.append(row)
                    row = []

                meter_data_lines_ctr += 1

        self.bepu_data = data

    def parse_ls_b(self):
        skipline_substrings = [
            "------",
            "======",
            "*",
            "MULTIPLIER",
            "FLOOR  AREA",
            "VOLUME",
            "TIME",
            "DRY-BULB TEMP",
            "WET-BULB TEMP",
            "TOT HORIZONTAL SOLAR RAD",
            "WINDSPEED AT SPACE",
            "CLOUD AMOUNT 0(CLEAR)-10",
            "COOLING  LOAD",
            "HEATING  LOAD",
            "SENSIBLE",
            "(KBTU/H)    ( KW )  (KBTU/H)  ( KW )",
        ]
        skipline_start_substrings = [
            "SPACE",
        ]
        headers = [
            "SPACE NAME",
            "LOAD CATEGORY",
            "COOLING SENSIBLE (KBTU/H)",
            "COOLING LATENT (KBTU/H)",
            "HEATING SENSIBLE (KBTU/H)",
            "TOTAL COOLING INTENSITY (BTU/H/FT2)",
            "TOTAL HEATING INTENSITY (BTU/H/FT2)",
        ]
        category_name_span = (5, 28)
        category_spans = [(28, 36), (48, 56), (85, 93)]
        total_spans = [(28, 36), (75, 83)]
        data = [headers]
        lines = self.report_contents['LS-B']

        for space_name, line in lines:

            if (
                any(substring in line for substring in skipline_substrings)
                or any(line.startswith(skipline_start_substring) for skipline_start_substring in skipline_start_substrings)
                or len(line.strip()) == 0
            ):
                continue

            load_category = line[category_name_span[0]:category_name_span[1]].strip()
            if "TOTAL LOAD" in load_category:
                line_data = [space_name, load_category, "", "", ""]
                for start, end in total_spans:
                    try:
                        line_data.append(self.clean(line[start:end].strip()))
                    except ValueError:
                        line_data.append("")

            else:
                line_data = [space_name, load_category]
                for start, end in category_spans:
                    try:
                        line_data.append(self.clean(line[start:end].strip()))
                    except ValueError:
                        line_data.append("")

            data.append(line_data)

        self.ls_b_data = data
        return

    def parse_lv_b(self):
        skipline_substrings = [
            "------",
            "BDL RUN",
        ]
        summaryline_substrings = [
            "CONDITIONED FLOOR AREA",
            "TOTAL INSTALLED LIGHTING POWER",
            "TOTAL INSTALLED EQUIPMENT POWER",
        ]
        segments = [(0, 37), (37, 43), (43, 47), (47, 54), (54, 62), (62, 69), (69, 77), (77, 91), (91, 97), (97, 109),
                    (109, 121)]
        summary_segments = [(0, 32), (35, 48), (48, 52)]
        headers = [
            "Floor Name",
            "Space Name",
            "Multiplier",
            "Space Type",
            "Azimuth",
            "LPD",
            "People",
            "EPD",
            "Infil. Method",
            "ACH",
            "Area",
            "Volume",
        ]
        data = [headers]
        lines = self.report_contents['LV-B']

        floor = None
        for line in lines:

            if any(substring in line for substring in skipline_substrings) or len(line.strip()) == 0:
                continue

            if "BUILDING TOTALS" in line:
                line_array = ["Summary"]
                for start, end in segments:
                    try:
                        line_array.append(line[start:end].strip())
                    except ValueError:
                        line_array.append("")
                data.append(line_array)

            elif any(substring in line for substring in summaryline_substrings):
                line_array = ["Summary"]
                for start, end in summary_segments:
                    try:
                        line_array.append(line[start:end].strip())
                    except ValueError:
                        line_array.append("")
                data.append(line_array)

            elif "Spaces on floor:" in line:
                floor = line.split("Spaces on floor:")[1].strip()

            elif floor is not None:
                line_array = [floor]
                for start, end in segments:
                    try:
                        line_array.append(line[start:end].strip())
                    except ValueError:
                        line_array.append("")
                data.append(line_array)

        self.lv_b_data = data
        return

    def parse_lv_d(self):
        skipline_substrings = [
            "------",
            "BDL RUN",
            "NUMBER OF EXTERIOR",
            "U-VALUE INCLUDES OUTSIDE FILM",
            "- - - W I N D O W S - - -",
            "SURFACE                                U-VALUE",
            "(BTU/HR-SQFT-F)     (SQFT)  (BTU/HR-SQFT-F)",
            "U-VALUE/WINDOWS      U-VALUE/WALLS",
            "(BTU/HR-SQFT-F)     (BTU/HR-SQFT-F)"
        ]

        summary_start = "AVERAGE             AVERAGE         AVERAGE U-VALUE         WINDOW         WALL           WINDOW+WALL"
        summary_started = False

        segments = [(0, 41), (41, 48), (48, 62), (62, 75), (75, 89), (89, 100), (100, 118), (118, 128)]
        summary_segments = [(0, 20), (20, 30), (30, 50), (50, 70), (70, 90), (90, 105), (105, 121)]
        data = [["", "", "", "", "", "", "", "", ""],
                ["Space Name", "Surface Name", "U-Value", "Area", "U-Value", "Area", "U-Value", "Area", "Azimuth"]]
        summary_data = [[]]

        lines = self.report_contents['LV-D']
        line_array = []

        for line in lines:
            if any(substring in line for substring in skipline_substrings) or len(line.strip()) == 0:
                continue

            if summary_started:
                summ_line_array = []
                for start, end in summary_segments:
                    try:
                        summ_line_array.append(line[start:end].strip())
                    except ValueError:
                        summ_line_array.append("")
                summary_data.append(summ_line_array)

            elif summary_start in line:
                summary_started = True

            else:
                if "in space:" not in line:
                    line_array = []
                    for start, end in segments:
                        try:
                            line_array.append(line[start:end].strip())
                        except ValueError:
                            line_array.append("")
                else:
                    line_array.insert(0, line.split("in space:")[1].strip())
                    data.append(line_array)

        self.lv_d_data = (data, summary_data)
        return

    def parse_pv_a(self):
        skipline_substrings = [
            "------",
            "BDL RUN",
            "HEATING     COOLING      LOOP",
            "DEMAND      DEMAND       FLOW",
            "(MBTU/HR)   (MBTU/HR)     (GPM)",
            "FLOW        HEAD      SETPOINT",
            "ATTACHED TO                (GPM)       ( FT)",
            "CAPACITY      FLOW        HEAD",
            "CAPACITY      FLOW         EIR",
            "CAPACITY      FLOW        EIR",
            "(MBTU/HR)     (GPM)       ( FT)",
            "(MBTU/HR) (GAL/MIN   )    (FRAC)",
            "(MBTU/HR) (GAL/MIN   )  (FRAC)",
            "*** DW-HEATERS ***",
            ]

        loop_start = "*** CIRCULATION LOOPS ***"
        loop_data_started = False
        pump_start = "*** PUMPS ***"
        pump_data_started = False
        prim_start = "*** PRIMARY EQUIPMENT ***"
        prim_data_started = False

        loop_segments = [(0, 13), (13, 25), (25, 37), (37, 49), (49, 61), (61, 73), (73, 86),
                         (86, 97), (97, 109), (109, 120)]
        pump_segments = [(0, 40), (40, 48), (48, 60), (60, 72), (72, 84), (84, 96), (96, 108),
                         (108, 119)]

        prim_segments = []
        if self.doe_version == "DOE-2.2":
            prim_segments = [(0, 19), (19, 53), (53, 65), (65, 77), (77, 89), (89, 101), (101, 112)]
        elif self.doe_version == "DOE-2.3":
            prim_segments = [(0, 19), (19, 53), (53, 65), (65, 77), (77, 88)]

        lines = self.report_contents['PV-A']

        loop_data = [
            ["Loop Name", "Heat Demand \n(MMBtu/h)", "Cool Demand \n(MMBtu/h)", "Loop Flow \n(GPM)", "Total Head \n(ft)", "Supply UA Product \n(Btu/h-F)", "Supply Loss DT \n(F)", "Return UA Product \n(Btu/h-F)",
             "Return Loss DT \n(F)", "Loop Volume \n(Gal)", "Fluid Heat Cap. \n(Btu/lb-F)"]]
        pump_data = [
            ["Pump Name", "Qty", "Attached To", "Attached Eqp Type", "Flow \n(GPM)", "Head \n(ft)", "Head Setpoint \n(ft)", "Capacity Control", "Power \n(kW)", "Mech. Eff.", "Motor Eff."]]
        prim_data = [
            ["Equipment Name", "Equipment Type", "Attached To", "Capacity \n(MMBtu/h)", "Flow \n(GPM)", "Head \n(ft)"]]

        system_line = None
        previous_line = None
        for line in lines:
            if any(substring in line for substring in skipline_substrings) or len(line.strip()) == 0:
                previous_line = line
                continue

            if loop_start in line:
                loop_data_started = True
                prim_data_started = False
                previous_line = line
                continue

            elif pump_start in line:
                pump_data_started = True
                loop_data_started = False
                previous_line = line
                continue

            elif prim_start in line:
                prim_data_started = True
                loop_data_started = False
                pump_data_started = False
                previous_line = line
                continue

            if loop_data_started:
                if previous_line is not None:
                    loop_line_array = [previous_line[0:32]]
                else:
                    loop_line_array = []

                for start, end in loop_segments:
                    try:
                        loop_line_array.append(line[start:end].strip())
                    except ValueError:
                        loop_line_array.append("")
                if all(segment == "" for segment in loop_line_array[-6:]):
                    previous_line = line
                    continue
                loop_data.append(loop_line_array)

            elif pump_data_started:
                pump_line_array = [previous_line[0:32], ""]
                for start, end in pump_segments:
                    try:
                        pump_line_array.append(line[start:end].strip())
                    except ValueError:
                        pump_line_array.append("")
                if all(segment == "" for segment in pump_line_array[-6:]):
                    previous_line = line
                    continue
                pump_line_array.insert(3, "")
                pump_data.append(pump_line_array)

            elif prim_data_started:
                prim_line_array = []
                for start, end in prim_segments:
                    try:
                        prim_line_array.append(line[start:end].strip())
                    except ValueError:
                        prim_line_array.append("")
                if all(segment == "" for segment in prim_line_array[-3:]):
                    system_line = line
                    continue
                if prim_line_array[0] == "" and all(segment != "" for segment in prim_line_array[1:5]):
                    start, end = prim_segments[0]
                    prim_line_array[0] = previous_line[start:end].strip()
                    start, end = prim_segments[1]
                    prim_line_array[1] = previous_line[start:end].strip()

                prim_line_array.insert(0, system_line[0:32])
                prim_data.append(prim_line_array)

            previous_line = line

        self.pv_a_data = (loop_data, pump_data, prim_data)
        return

    def parse_ps_c(self):
        skipline_substrings = [
            "------",
            "BDL RUN",
            "(MBTU)      (MBTU)",
            "(KBTU/HR)   (KBTU/HR)"
        ]
        lines = self.report_contents.get("PS-C", [])
        data = [["System", "Type", "Cool Load (MBTU)", "Heat Load (MBTU)", "Elec Use (kWh)", "Fuel Use (MBTU)", "Data Type", "PLR 0_10", "PLR 10_20", "PLR 20_30", "PLR 30_40", "PLR 40_50", "PLR 50_60", "PLR 60_70", "PLR 70_80", "PLR 80_90", "PLR 90_100", "PLR 100+", "Total Run Hours"]]
        current_system = None

        for line in lines:
            if any(substring in line for substring in skipline_substrings) or len(line.strip()) == 0:
                continue

            stripped = line.strip()

            if not stripped.startswith(("SUM", "PEAK", "MON/DAY")):
                current_system = stripped
                continue

            entry_type = stripped.split()[0]
            cool_val = line[11:21].strip()
            heat_val = line[23:33].strip()
            elec_val = line[35:45].strip()
            fuel_val = line[47:57].strip()
            data_type = line[58:62]
            plr_values = [line[i:i+6].strip() for i in range(62, 132, 6)]

            data.append([
                current_system,
                entry_type,
                self.clean(cool_val),
                self.clean(heat_val),
                self.clean(elec_val),
                self.clean(fuel_val),
                data_type,
                *[self.clean(plr) for plr in plr_values],
            ])

        self.ps_c_data = data

    def parse_ps_h(self):
        return

    def parse_ss_a(self):
        skipline_substrings = [
            "------",
            "BDL RUN",
            "- - - - - - - - C O O L I N G - - - - - - - -",
            "MAXIMUM         ELEC-    MAXIMUM",
            "COOLING     TIME   DRY-  WET-",
            "ENERGY   OF MAX   BULB  BULB",
            "MONTH     (MBTU)   DY  HR   TEMP  TEMP",
        ]
        segments = [(0, 5), (5, 16), (16, 22), (22, 25), (25, 32), (32, 38), (38, 52), (52, 67), (67, 72), (72, 76),
                    (76, 83), (83, 89), (89, 103), (103, 117), (117, 128)]
        lines = self.report_contents['SS-A']  # Exclude the last line
        data = [
            ["System", "Month", "Cooling Energy \n(MBTU)", "Peak Cooling Day", "Peak Cooling Hour", "Dry Bulb Temp",
             "Wet Bulb Temp", "Max Cooling Load \n(KBTU/H)", "Heating Energy \n(MBTU)", "Peak Heating Day",
             "Peak Heating Hour", "Dry Bulb Temp", "Wet Bulb Temp", "Max Heating Load \n(KBTU/H)",
             "Electrical Energy \n(KWH)", "Peak Electrical Load \n(KW)"]]

        for line in lines:
            if any(substring in line[1] for substring in skipline_substrings) or len(line[1].strip()) == 0:
                continue

            line_array = [line[0]]
            for start, end in segments:
                line_array.append(line[1][start:end].strip())
            data.append(line_array)
        self.ss_a_data = data
        return

    def parse_ss_b(self):
        skipline_substrings = [
            "------",
            "BDL RUN",
        ]
        segments = [(0, 5), (5, 16), (16, 22), (22, 25), (25, 32), (32, 38), (38, 52), (52, 67), (67, 72), (72, 76),
                    (76, 83), (83, 89), (89, 103), (103, 117), (117, 128)]
        lines = self.report_contents['SS-B']
        data = [
            ["", "", ""]]

        for line in lines:
            if any(substring in line[1] for substring in skipline_substrings) or len(line[1].strip()) == 0:
                continue

            line_array = []
            for start, end in segments:
                line_array.append(line[1][start:end].strip())
            data.append(line_array)
        self.ss_b_data = data
        return

    def parse_ss_f(self):
        skipline_substrings = [
            "------",
            "BDL RUN",
            '- - - -',
            "HEAT          HEAT",
            "EXTRACTION      ADDITION",
            "ENERGY        ENERGY",
            "(MBTU)        (MBTU)"
        ]
        lines = self.report_contents.get("SS-F", [])
        data = [["Zone", "Month", "Heat Extraction (MBTU)", "Heat Addition (MBTU)",
                 "Baseboard Energy (MBTU)", "Max Baseboard Load (kBTU/hr)",
                 "Max Zone Temp (°F)", "Min Zone Temp (°F)",
                 "Hours Under Heated", "Hours Under Cooled"]]

        ssf_month_row = re.compile(r"^(JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC)\s+")

        for idx, (zone_name, line) in enumerate(lines):

            if any(substring in line for substring in skipline_substrings) or len(line.strip()) == 0:
                continue

            stripped = line.strip()

            if ssf_month_row.match(stripped):
                match = re.match(
                    r"^(JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC)\s+"
                    r"([-\d\.Ee]+)\s+([-\d\.Ee]+)\s+([-\d\.Ee]+)\s+([-\d\.Ee]+)\s+"
                    r"([-\d\.Ee]+)\s+([-\d\.Ee]+)\s+([-\d\.Ee]+)\s+([-\d\.Ee]+)",
                    stripped
                )
                if match:
                    month, heat_ex, heat_add, base_energy, base_load, max_temp, min_temp, hrs_heat, hrs_cool = match.groups()
                    data.append([
                        zone_name,
                        month,
                        self.clean(heat_ex),
                        self.clean(heat_add),
                        self.clean(base_energy),
                        self.clean(base_load),
                        self.clean(max_temp),
                        self.clean(min_temp),
                        self.clean(hrs_heat),
                        self.clean(hrs_cool)
                    ])

        self.ss_f_data = data

    def parse_ss_g(self):
        skipline_substrings = [
            "------",
            "BDL RUN",
            "- - - - - - - - C O O L I N G - - - - - - - -",
            "MAXIMUM         ELEC-    MAXIMUM",
            "COOLING     TIME   DRY-  WET-",
            "ENERGY   OF MAX   BULB  BULB",
            "MONTH     (MBTU)   DY  HR   TEMP  TEMP",
        ]
        segments = [(0, 5), (5, 16), (16, 22), (22, 25), (25, 32), (32, 38), (38, 52), (52, 67), (67, 72), (72, 76),
                    (76, 83), (83, 89), (89, 103), (103, 117), (117, 128)]
        lines = self.report_contents['SS-G']
        data = [
            ["System", "Month", "Cooling Energy \n(MBTU)", "Peak Cooling Day", "Peak Cooling Hour", "Dry Bulb Temp",
             "Wet Bulb Temp", "Max Cooling Load \n(KBTU/H)", "Heating Energy \n(MBTU)", "Peak Heating Day",
             "Peak Heating Hour", "Dry Bulb Temp", "Wet Bulb Temp", "Max Heating Load \n(KBTU/H)",
             "Electrical Energy \n(KWH)", "Peak Electrical Load \n(KW)"]]

        for line in lines:
            if any(substring in line[1] for substring in skipline_substrings) or len(line[1].strip()) == 0:
                continue

            line_array = [line[0]]
            for start, end in segments:
                line_array.append(line[1][start:end].strip())
            data.append(line_array)
        self.ss_g_data = data
        return

    def parse_ss_h(self):
        skipline_substrings = [
            "------",
            "BDL RUN",
            "- - - - -",
            "- -F A N   E L E C- - -",
            "MAXIMUM                   MAXIMUM",
            "FAN         FAN",
            "ENERGY        LOAD",
            "(KWH)        (KW)"
        ]
        lines = self.report_contents.get("SS-H", [])
        data = [["System", "Month", "Fan Electric Energy (kWh)", "Maximum Fan Load (kW)",
                 "Gas Heat Energy (MBtu)", "Maximum Gas Heat Load (kBtu/hr)",
                 "Gas Cool Energy (MBtu)", "Maximum Gas Cool Load (kBtu/hr)",
                 "Electric Heat Energy (kWh)", "Maximum Electric Heat Load (kW)",
                 "Electric Cool Energy (kWh)", "Maximum Electric Cool Load (kW)"]]

        ss_h_row = re.compile(r"^(JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC|TOTAL|MAX)\s")

        for system_name, line in lines:

            if any(substring in line for substring in skipline_substrings) or len(line.strip()) == 0:
                continue

            stripped = line.strip()
            if not stripped:
                continue

            if ss_h_row.match(stripped):
                parts = stripped.split()
                month_tag = parts[0].upper()
                values = parts[1:]
                row = [system_name, month_tag] + [""] * 10

                if month_tag == "TOTAL":
                    row[2] = self.clean(values[0])  # Fan Electric Energy
                    row[4] = self.clean(values[1])  # Gas Heat Energy
                    row[6] = self.clean(values[2])  # Gas Cool Energy
                    row[8] = self.clean(values[3])  # Electric Heat Energy
                    row[10] = self.clean(values[4])  # Electric Cool Energy
                elif month_tag == "MAX":
                    row[3] = self.clean(values[0])  # Max Fan Load
                    row[5] = self.clean(values[1])  # Max Gas Heat Load
                    row[7] = self.clean(values[2])  # Max Gas Cool Load
                    row[9] = self.clean(values[3])  # Max Elec Heat Load
                    row[11] = self.clean(values[4])  # Max Elec Cool Load
                else:
                    row[2:12] = values[:10]

                data.append(row)

        self.ss_h_data = data

    def parse_ss_l(self):
        previous_line = ""
        lines = self.report_contents.get("SS-L", [])
        if not lines:
            self.ss_l_data = None
            return

        month_line_regex = re.compile(r"^(JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC|ANNUAL)\b", re.IGNORECASE)
        breakdown_start_regex = re.compile(r"BREAKDOWN OF ANNUAL FAN POWER USAGE", re.IGNORECASE)

        columns = [
            "System", "Month", "FAN ELEC Heating", "FAN ELEC Cooling", "FAN ELEC Heat & Cool",
            "FAN ELEC Floating", "PLR_00_10", "PLR_10_20", "PLR_20_30", "PLR_30_40",
            "PLR_40_50", "PLR_50_60", "PLR_60_70", "PLR_70_80", "PLR_80_90", "PLR_90_100",
            "PLR_100+", "Total Run Hours", "Annual FAN ELEC (kWh)"
        ]

        data = [columns]
        in_month_section = False
        in_breakdown_section = False
        breakdown_headers_line = False
        breakdown_dashes_line = False

        for system_name, line in lines:
            stripped = line.strip()

            # Detect dashed line to begin monthly data capture
            if "---" in stripped and not in_month_section:
                in_month_section = True
                continue

            # Capture monthly data
            if in_month_section and month_line_regex.match(stripped):
                parts = stripped.split()
                if len(parts) >= 17:
                    row = [system_name] + parts + [""]
                    if len(row) == len(columns):
                        data.append(row)
                continue

            # Detect start of breakdown section
            if breakdown_start_regex.search(stripped):
                in_breakdown_section = True
                breakdown_dashes_line = None
                breakdown_headers_line = None
                continue

            # Save the dashed line and header for parsing spans
            if in_breakdown_section and not breakdown_dashes_line and "-" in stripped:
                breakdown_dashes_line = stripped
                continue
            if in_breakdown_section and breakdown_dashes_line and not breakdown_headers_line:
                breakdown_headers_line = previous_line.strip()  # previous_line holds the line before dashed one
                continue

            # Extract fixed-width values using spans from the dashed line
            if in_breakdown_section and breakdown_dashes_line and breakdown_headers_line:
                spans = [m.span() for m in re.finditer(r"-{3,}", breakdown_dashes_line)]
                col_names = [breakdown_headers_line[start:end].strip().upper() for start, end in spans]
                col_names = [c + " (KWH)" if not c.endswith("(KWH)") else c for c in col_names]
                full_columns = ["System", "Month"] + col_names

                parts = [line.strip()[start:end] for start, end in spans]
                rows = [
                    [system_name, "BREAKDOWN", "SUPPLY (KWH)", "HOT DECK (KWH)", "RETURN (KWH)", "RELIEF (KWH)", "PIU TERMINALS (KWH)", "ZONE EXH (KWH)", "TOTAL (KWH)"],
                    [system_name, "BREAKDOWN"] + [self.clean(v) for v in parts]
                ]
                if len(rows[1]) == len(full_columns):
                    data.extend(rows)

                # Reset for next system
                in_month_section = False
                in_breakdown_section = False
                breakdown_dashes_line = None
                breakdown_headers_line = None
                continue

            previous_line = line.strip()

        self.ss_l_data = data if len(data) > 1 else None

    def parse_ss_r(self):
        segments = [(18, 26), (27, 35), (36, 44), (45, 53), (59, 63), (65, 69), (71, 75),
                    (77, 81), (83, 87), (89, 93), (95, 99), (101, 105), (107, 111),
                    (113, 117), (119, 123), (125, 129)]
        lines = self.report_contents['SS-R']
        data = [
            ["System Name", "Zone Name", "Max. Heat Hours", "Max. Cool Hours", "Unmet Heat Hours", "Unmet Cool Hours",
             "00-10", "10-20", "20-30", "30-40",
             "40-50", "50-60", "60-70", "70-80", "80-90", "90-100", "100+", "Run Hours"]]

        previous_line = None
        for line in lines:
            # Check if the line contains only numbers (using regular expression)
            if re.match(r'^[\d\s]+$', line[1].strip()):
                if previous_line is not None:
                    line_array = [line[0], previous_line[1].strip()]
                    for start, end in segments:
                        try:
                            line_array.append(int(line[1][start:end].strip()))
                        except ValueError:
                            line_array.append("")
                    data.append(line_array)
            previous_line = line

        self.ss_r_data = data
        return

    def parse_sv_a(self):
        skipline_substrings = [
            "------",
            "*** ",
            "BDL RUN",
            "SYSTEM   ALTITUDE       AREA",
            "TYPE     FACTOR    (SQFT )",
            "FAN   CAPACITY     FACTOR",
            "TYPE    (CFM )     (FRAC)",
            "ZONE                       FLOW      FLOW",
            "NAME                     (CFM )    (CFM )",
            "(BASEBOARDS)",
            "MIXED AIR      ZONE",
            "(CFM )     (CFM )    MULT",
            "VRF BRANCH GAS PIPE NOMINAL DIA"
        ]

        sys_start = "FLOOR               OUTSIDE    COOLING"
        sys_data_started = False
        fan_start = "DIVERSITY    POWER       FAN"
        fan_data_started = False
        zn_start = "SUPPLY   EXHAUST             MINIMUM"
        zn_data_started = False
        doas_start = "-----  OA ATTACHED TO  -----"
        doas_data_started = False

        sys_segments = [(0, 13), (13, 19), (19, 30), (30, 42), (42, 52), (52, 63), (63, 74),
                        (74, 85), (85, 96), (96, 107), (107, 118)]
        fan_segments = [(0, 9), (9, 20), (20, 31), (31, 40), (40, 50), (50, 61), (61, 69),
                        (69, 77), (77, 89), (89, 99), (99, 109), (109, 118)]
        zn_segments = [(0, 28), (28, 37), (37, 47), (47, 57), (57, 67), (67, 77), (77, 87),
                       (87, 97), (97, 107), (107, 117), (117, 127), (127, 131)]
        doas_segments = [(0, 37), (37, 48), (48, 59), (59, 67)]

        lines = self.report_contents['SV-A']
        sys_data = [
            ["System Name", "Type", "Alt. Factor", "Floor Area", "Max Occ.", "OA Ratio", "Cooling \n(kBtu/h)", "SHR",
             "Heating \n(kBtu/h)", "Cool EIR", "Heat EIR", "HP Supp. \n(kBtu/h)"]]
        fan_data = [
            ["System Name", "Type", "Flow Cap. \n(CFM)", "Div. Factor", "Demand \n(kW)", "Fan dT \n(F)",
             "SP \n(in. H2O)",
             "Total Eff.", "Mech. Eff.", "Fan Placement", "Fan Control", "Max. Fan \nRatio", "Min. Fan \nRatio"]]
        zn_data = [
            ["System Name", "Zone Name", "Supply \n(CFM)", "Exhaust \n(CFM)", "Fan \n(kW)", "Min. Flow \nRatio",
             "OA \n(CFM)",
             "Cooling \n(kBtu/h)", "SHR", "Extr. \n(kBtu/h)", "Heating \n(kBtu/h)", "Addition \n(kBtu/h)", "Zn Mult."]]
        doas_data = [
            ["System Name", "Zone Name", "Mixed Air \n(CFM)", "Zone \n(CFM)", "Mult."],
        ]

        for line in lines:
            if any(substring in line[1] for substring in skipline_substrings) or len(line[1].strip()) == 0:
                continue

            if sys_start in line[1]:
                sys_data_started = True
                zn_data_started = False
                continue

            elif fan_start in line[1]:
                fan_data_started = True
                sys_data_started = False
                continue

            elif zn_start in line[1]:
                zn_data_started = True
                sys_data_started = False
                fan_data_started = False
                continue

            elif doas_start in line[1]:
                doas_data_started = True
                sys_data_started = False
                fan_data_started = False
                zn_data_started = False
                continue

            if sys_data_started:
                sys_line_array = [line[0]]
                for start, end in sys_segments:
                    try:
                        sys_line_array.append(line[1][start:end].strip())
                    except ValueError:
                        sys_line_array.append("")
                sys_data.append(sys_line_array)

            elif fan_data_started:
                fan_line_array = [line[0]]
                for start, end in fan_segments:
                    try:
                        fan_line_array.append(line[1][start:end].strip())
                    except ValueError:
                        fan_line_array.append("")
                fan_data.append(fan_line_array)

            elif zn_data_started:
                zn_line_array = [line[0]]
                for start, end in zn_segments:
                    try:
                        zn_line_array.append(line[1][start:end].strip())
                    except ValueError:
                        zn_line_array.append("")
                zn_data.append(zn_line_array)

            elif doas_data_started:
                doas_line_array = [line[0]]
                for start, end in doas_segments:
                    try:
                        doas_line_array.append(line[1][start:end].strip())
                    except ValueError:
                        doas_line_array.append("")
                doas_data.append(doas_line_array)

        self.sv_a_data = (sys_data, fan_data, zn_data, doas_data)
        return

    def write_excel(self):
        workbook = xlsxwriter.Workbook(self.wb_name, {'nan_inf_to_errors': True})

        # --- Formats
        header_format = workbook.add_format({
            'bold': True,
            'align': 'center',
            'valign': 'vcenter',
            'font_size': 14,
            'text_wrap': True,
            'bg_color': '#D9E1F2',
        })
        caution_format = workbook.add_format({
            'font_color': 'red',
            'bold': True,
            'text_wrap': True
        })
        number_format = workbook.add_format({'num_format': '0.00'})
        string_format = workbook.add_format({'num_format': '@'})
        bold_format = workbook.add_format({'bold': True})
        caution_merge_fmt = workbook.add_format({
            'bold': True, 'font_color': '#FF0000', 'text_wrap': True,
            'align': 'left', 'valign': 'vbottom'
        })

        # === BEPU ===
        if self.bepu_data:
            bepu_ws = workbook.add_worksheet("BEPU")
            for row, data in enumerate(self.bepu_data):
                if row == 0:
                    bepu_ws.write_row(row, 0, data, header_format)
                else:
                    data = try_convert_element_to_float(data)
                    bepu_ws.write_row(row, 0, data)
            bepu_ws.set_column(0, len(self.bepu_data[0]) - 1, 15)

        # === LS-B ===
        if self.ls_b_data:
            ls_b_ws = workbook.add_worksheet('LS-B')
            numeric_cols = {2, 3, 4, 5, 6}  # columns containing numeric data

            for row_idx, row_data in enumerate(self.ls_b_data):
                for col_idx, value in enumerate(row_data):
                    if row_idx == 0:
                        ls_b_ws.write(row_idx, col_idx, value, header_format)
                    else:
                        if col_idx in numeric_cols:
                            try:
                                ls_b_ws.write_number(row_idx, col_idx, float(value), number_format)
                            except ValueError:
                                ls_b_ws.write(row_idx, col_idx, value, string_format)
                        else:
                            ls_b_ws.write_string(row_idx, col_idx, str(value), string_format)

            # Column width tuning for readability
            ls_b_ws.set_column(0, 0, 30)  # SPACE NAME
            ls_b_ws.set_column(1, 1, 27)  # LOAD CATEGORY
            ls_b_ws.set_column(2, 4, 18)  # Cooling/Heating sensible
            ls_b_ws.set_column(5, 6, 22)  # Intensities

            # Enable autofilter over the entire table (no default filters applied)
            ls_b_ws.autofilter(0, 0, len(self.ls_b_data) - 1, len(self.ls_b_data[0]) - 1)

        # === LV-B ===
        if self.lv_b_data:
            lv_b_ws = workbook.add_worksheet('LV-B')
            numeric_cols = {2, 4, 5, 6, 7, 9, 10, 11}
            for row_idx, row_data in enumerate(self.lv_b_data):
                for col_idx, value in enumerate(row_data):
                    if row_idx == 0:
                        lv_b_ws.write(row_idx, col_idx, value, header_format)
                    else:
                        if col_idx in numeric_cols:
                            try:
                                lv_b_ws.write_number(row_idx, col_idx, float(value), number_format)
                            except ValueError:
                                lv_b_ws.write(row_idx, col_idx, value, string_format)
                        else:
                            lv_b_ws.write_string(row_idx, col_idx, str(value), string_format)
            lv_b_ws.set_column(0, 0, 19.94)
            lv_b_ws.set_column(1, 1, 32.04)
            lv_b_ws.set_column(2, 5, 13.57)
            lv_b_ws.set_column(5, 7, 11.39)
            lv_b_ws.set_column(8, 8, 16.43)
            lv_b_ws.set_column(9, 11, 11.39)

        # === LV-D ===
        if self.lv_d_data:
            lv_d_ws = workbook.add_worksheet('LV-D')
            for row, data in enumerate(self.lv_d_data[0]):
                if row == 1:
                    lv_d_ws.write_row(row, 0, data, header_format)
                else:
                    data = try_convert_element_to_float(data)
                    lv_d_ws.write_row(row, 0, data)
            lv_d_ws.set_column(0, 0, 31.14)
            lv_d_ws.set_column(1, 1, 38.71)
            lv_d_ws.set_column(2, 8, 12.14)
            lv_d_ws.merge_range("C1:D1", "Windows", header_format)
            lv_d_ws.merge_range("E1:F1", "Walls", header_format)
            lv_d_ws.merge_range("G1:H1", "Walls+Windows", header_format)

        # === PS-C ===
        if self.ps_c_data:
            ps_c_ws = workbook.add_worksheet("PS-C")
            t_column_format = workbook.add_format({
                'bold': True,
                'font_size': 14,
                'text_wrap': True,
                'bg_color': '#D9E1F2'
            })
            caution_format_ps_c = workbook.add_format({'font_color': 'red', 'bold': True, 'text_wrap': False})

            for row, data in enumerate(self.ps_c_data):
                if row == 0:
                    ps_c_ws.write_row(row, 0, data, header_format)
                else:
                    data = try_convert_element_to_float(data)
                    ps_c_ws.write_row(row, 0, data)
            ps_c_ws.set_column(0, 0, 25)
            ps_c_ws.set_column(1, 1, 10)
            ps_c_ws.set_column(2, 5, 16)

            ps_c_ws.autofilter(0, 0, len(self.ps_c_data) - 1, len(self.ps_c_data[0]) - 1)
            ps_c_ws.filter_column(1, 'Type == SUM')  # keep your filter UI
            ps_c_ws.set_column('G:S', None, None, {'hidden': True})  # Hide G:S

            for row_idx in range(1, len(self.ps_c_data)):
                excel_row = row_idx + 1
                formula = f'=IFERROR(ABS(D{excel_row}/((E{excel_row}*3.412)/1000)),"")'
                ps_c_ws.write_formula(row_idx, 19, formula)  # T column
            ps_c_ws.write('T1', 'Heating Efficiency', t_column_format)
            caution_text = ("⚠️ If a piece of equipment provides both heating and cooling, "
                            "the calculated heating efficiency will not be accurate.")
            ps_c_ws.write('U1', caution_text, caution_format_ps_c)

            # Hide rows not containing 'SUM'
            # Header row is 0; start hiding from Excel row 2 (0-based row 1)
            for r in range(1, len(self.ps_c_data)):
                row_vals = self.ps_c_data[r]
                if not any(isinstance(v, str) and 'SUM' in v for v in row_vals):
                    ps_c_ws.set_row(r, options={'hidden': True})

        # === PV-A ===
        if self.pv_a_data:
            pv_a0_ws = workbook.add_worksheet('PV-A Loops')
            for row, data in enumerate(self.pv_a_data[0]):
                if row == 0:
                    pv_a0_ws.write_row(row, 0, data, header_format)
                else:
                    data = try_convert_element_to_float(data)
                    pv_a0_ws.write_row(row, 0, data)
            pv_a0_ws.set_column(0, 0, 27.86)
            pv_a0_ws.set_column(1, 10, 13.57)

            pv_a1_ws = workbook.add_worksheet('PV-A Pumps')
            for row, data in enumerate(self.pv_a_data[1]):
                if row == 0:
                    pv_a1_ws.write_row(row, 0, data, header_format)
                else:
                    data = try_convert_element_to_float(data)
                    pv_a1_ws.write_row(row, 0, data)
            pv_a1_ws.set_column(0, 0, 27.86)
            pv_a1_ws.set_column(1, 1, 10)
            pv_a1_ws.set_column(2, 2, 24.29)
            pv_a1_ws.set_column(3, 4, 15)
            pv_a1_ws.set_column(5, 10, 13.57)

            pv_a2_ws = workbook.add_worksheet('PV-A Equip.')
            for row, data in enumerate(self.pv_a_data[2]):
                if row == 0:
                    pv_a2_ws.write_row(row, 0, data, header_format)
                else:
                    data = try_convert_element_to_float(data)
                    pv_a2_ws.write_row(row, 0, data)
            pv_a2_ws.set_column(0, 3, 27.86)
            pv_a2_ws.set_column(3, 5, 13.57)

        # === SS-A ===
        if self.ss_a_data:
            ss_a_ws = workbook.add_worksheet('SS-A')
            calcs_heading_format = workbook.add_format({'bold': True, 'text_wrap': True})

            # Cautions
            ss_a_ws.write('A1', "⚠️ Load and Efficiency will not be Accurate unless the 'TOTAL' Filter is Applied.", caution_format)
            ss_a_ws.write('B1', "⚠️ QC that cells D2 and J2 align with the BEPU tab and that the BEPU tab is not missing data.", caution_format)

            # Headings
            ss_a_ws.write('C1', "Total Cooling Load, MMBtu", calcs_heading_format)
            ss_a_ws.write('D1', "Cooling Consumption from BEPU tab, MMBtu", calcs_heading_format)
            ss_a_ws.write('E1', "Whole Building Annualized Cooling Efficiency, COP", calcs_heading_format)
            ss_a_ws.write('I1', "Total Heating Load, MMBtu", calcs_heading_format)
            ss_a_ws.write('J1', "Heating Consumption from BEPU tab, MMBtu", calcs_heading_format)
            ss_a_ws.write('K1', "Whole Building Annualized Heating Efficiency, COP", calcs_heading_format)

            # Data (shift down 2 rows like before)
            for row, data in enumerate(self.ss_a_data):
                excel_row = row + 2
                if row == 0:
                    ss_a_ws.write_row(excel_row, 0, data, header_format)
                else:
                    data = try_convert_element_to_float(data)
                    ss_a_ws.write_row(excel_row, 0, data)
            ss_a_ws.set_column(0, 1, 18.71)
            ss_a_ws.set_column(2, 3, 9.29)
            ss_a_ws.set_column(4, 5, 11.00)
            ss_a_ws.set_column(6, 6, 12.14)
            ss_a_ws.set_column(7, 7, 18.71)
            ss_a_ws.set_column(8, 9, 9.29)
            ss_a_ws.set_column(10, 11, 11.00)
            ss_a_ws.set_column(12, 12, 12.14)
            ss_a_ws.set_column(13, 14, 11.57)

            header_row = 2  # Excel row 3
            ss_a_ws.autofilter(header_row, 0, len(self.ss_a_data) - 1, len(self.ss_a_data[0]) - 1)
            # Filter UI (label says Type in your original; value is in the "Month" column which can be TOTAL)
            ss_a_ws.filter_column(1, 'Type == TOTAL')

            # Totals / COP formulas on row 2
            start_row = 4
            end_row = len(self.ss_a_data) + 2
            ss_a_ws.write_formula('C2', f'=SUBTOTAL(9,C{start_row}:C{end_row})', bold_format)
            ss_a_ws.write_formula('D2', '=(BEPU!G2*3.412)/1000', bold_format)
            ss_a_ws.write_formula('E2', '=ABS(C2)/D2', bold_format)
            ss_a_ws.write_formula('I2', f'=SUBTOTAL(9,I{start_row}:I{end_row})', bold_format)
            ss_a_ws.write_formula('J2', '=(BEPU!F2+BEPU!L2)*3.412*(1/1000)', bold_format)
            ss_a_ws.write_formula('K2', '=ABS(I2)/J2', bold_format)

            # Hide all rows not containing 'TOTAL' (from Excel row 4 down)
            for r in range(3, len(self.ss_a_data) + 2):
                # map to data index: r-2
                row_vals = self.ss_a_data[r - 2]
                if not any(isinstance(v, str) and 'TOTAL' in v for v in row_vals):
                    ss_a_ws.set_row(r, options={'hidden': True})

        # === SS-F ===
        if self.ss_f_data:
            ss_f_ws = workbook.add_worksheet('SS-F')
            for row, data in enumerate(self.ss_f_data):
                if row == 0:
                    # Add a "System" column header at K1
                    hdr = data + ["System"]
                    ss_f_ws.write_row(row, 0, hdr, header_format)
                else:
                    data = try_convert_element_to_float(data)
                    ss_f_ws.write_row(row, 0, data)
            ss_f_ws.set_column(0, 1, 26)
            ss_f_ws.set_column(2, len(self.ss_f_data[0]) - 1, 14)

            # Caution (L1)
            ss_f_ws.set_column(11, 11, 40)
            total_bbrd_text = "Use Column E to Verify Baseboard is Actually being Modeled where Expected"
            ss_f_ws.write('L1', total_bbrd_text, caution_format)

            # Populate "System" formulas in column K (index 10), matching zone names to SS-R col B
            if self.ss_r_data:
                last_row_r = len(self.ss_r_data)  # includes header
                rng_a = f"'SS-R'!$A$2:$A${last_row_r}"
                rng_b = f"'SS-R'!$B$2:$B${last_row_r}"
                for r in range(1, len(self.ss_f_data)):
                    excel_row = r + 1
                    formula = f'=INDEX({rng_a},MATCH(A{excel_row},{rng_b},0),1)'
                    ss_f_ws.write_formula(r, 10, formula)  # K column

        # === SS-G ===
        if self.ss_g_data:
            ss_g_ws = workbook.add_worksheet('SS-G')
            for row, data in enumerate(self.ss_g_data):
                if row == 0:
                    ss_g_ws.write_row(row, 0, data, header_format)
                else:
                    data = try_convert_element_to_float(data)
                    ss_g_ws.write_row(row, 0, data)

        # === SS-H ===
        if self.ss_h_data:
            ss_h_ws = workbook.add_worksheet('SS-H')
            for row, data in enumerate(self.ss_h_data):
                if row == 0:
                    ss_h_ws.write_row(row, 0, data, header_format)
                else:
                    data = try_convert_element_to_float(data)
                    ss_h_ws.write_row(row, 0, data)
            ss_h_ws.set_column(0, 1, 20)
            ss_h_ws.set_column(2, len(self.ss_h_data[0]) - 1, 14)

            ss_h_ws.autofilter(0, 0, len(self.ss_h_data) - 1, len(self.ss_h_data[0]) - 1)
            ss_h_ws.filter_column(1, 'Type == TOTAL')  # Month column where 'TOTAL' appears
            ss_h_ws.set_column(12, 12, 40)

            # Hide rows not containing 'TOTAL' (from Excel row 2 down)
            for r in range(1, len(self.ss_h_data)):
                row_vals = self.ss_h_data[r]
                if not any(isinstance(v, str) and 'TOTAL' in v for v in row_vals):
                    ss_h_ws.set_row(r, options={'hidden': True})

        # === SS-L ===
        if self.ss_l_data:
            ss_l_ws = workbook.add_worksheet('SS-L')
            for row, data in enumerate(self.ss_l_data):
                if row == 0:
                    ss_l_ws.write_row(row, 0, data, header_format)
                else:
                    data = try_convert_element_to_float(data)
                    ss_l_ws.write_row(row, 0, data)
            ss_l_ws.set_column(0, 1, 20)
            ss_l_ws.set_column(2, len(self.ss_l_data[0]) - 1, 14)

            ss_l_ws.set_column('G:Q', None, None, {'hidden': True})
            ss_l_ws.autofilter(0, 0, len(self.ss_l_data) - 1, len(self.ss_l_data[0]) - 1)
            ss_l_ws.filter_column(1, 'Type == ANNUAL')
            ss_l_ws.set_column(20, 20, 40)

            # Hide rows not containing 'ANNUAL' (from Excel row 2 down)
            for r in range(1, len(self.ss_l_data)):
                row_vals = self.ss_l_data[r]
                if not any(isinstance(v, str) and 'ANNUAL' in v for v in row_vals):
                    ss_l_ws.set_row(r, options={'hidden': True})

        # === SS-R ===
        if self.ss_r_data:
            ss_r_ws = workbook.add_worksheet('SS-R')
            for row, data in enumerate(self.ss_r_data):
                if row == 0:
                    ss_r_ws.write_row(row, 0, data, header_format)
                else:
                    ss_r_ws.write_row(row, 0, data)
            ss_r_ws.set_column(0, 0, 27.86)
            ss_r_ws.set_column(1, 1, 38.57)
            ss_r_ws.set_column(2, 5, 14.57)
            ss_r_ws.set_column(17, 17, 15.14)

        # === SV-A ===
        if self.sv_a_data:
            # Systems
            sv_sys_ws = workbook.add_worksheet('SV-A Systems')
            for row, data in enumerate(self.sv_a_data[0]):
                if row == 0:
                    # Add extra headers M..P
                    hdr = data + ["Supply CFM", "OA CFM", "Supply CFM/sf", "OA CFM/sf"]
                    sv_sys_ws.write_row(row, 0, hdr, header_format)
                else:
                    data = try_convert_element_to_float(data)
                    sv_sys_ws.write_row(row, 0, data)
            sv_sys_ws.set_column(0, 0, 31.14)
            sv_sys_ws.set_column(1, 2, 13.57)
            sv_sys_ws.set_column(3, 3, 12.86)
            sv_sys_ws.set_column(4, 12, 11.43)

            # Fans
            sv_fan_ws = workbook.add_worksheet('SV-A Fans')
            for row, data in enumerate(self.sv_a_data[1]):
                if row == 0:
                    sv_fan_ws.write_row(row, 0, data, header_format)
                else:
                    data = try_convert_element_to_float(data)
                    sv_fan_ws.write_row(row, 0, data)
            sv_fan_ws.set_column(0, 0, 31.14)
            sv_fan_ws.set_column(1, 2, 13.57)
            sv_fan_ws.set_column(3, 3, 12.86)
            sv_fan_ws.set_column(4, 8, 11.43)
            sv_fan_ws.set_column(9, 9, 12.86)
            sv_fan_ws.set_column(10, 11, 11.43)

            # Zones
            sv_zn_ws = workbook.add_worksheet('SV-A Zones')
            for row, data in enumerate(self.sv_a_data[2]):
                if row == 0:
                    sv_zn_ws.write_row(row, 0, data, header_format)
                else:
                    data = try_convert_element_to_float(data)
                    sv_zn_ws.write_row(row, 0, data)
            sv_zn_ws.set_column(0, 0, 31.14)
            sv_zn_ws.set_column(1, 1, 26.43)
            sv_zn_ws.set_column(2, 2, 13.57)
            sv_zn_ws.set_column(3, 3, 12.86)
            sv_zn_ws.set_column(4, 11, 11.43)

            # DOAS (optional)
            if len(self.sv_a_data[3]) > 1:
                sv_doas_ws = workbook.add_worksheet('SV-A DOAS')
                for row, data in enumerate(self.sv_a_data[3]):
                    if row == 0:
                        sv_doas_ws.write_row(row, 0, data, header_format)
                    else:
                        data = try_convert_element_to_float(data)
                        sv_doas_ws.write_row(row, 0, data)

            # Add Systems formulas referencing Fans (M..P)
            last_fans_row = len(self.sv_a_data[1])  # includes header
            fans_a = f"'SV-A Fans'!$A$2:$A${last_fans_row}"
            fans_b = f"'SV-A Fans'!$B$2:$B${last_fans_row}"
            fans_c = f"'SV-A Fans'!$C$2:$C${last_fans_row}"
            for r in range(1, len(self.sv_a_data[0])):
                excel_row = r + 1
                # M (index 12): Supply CFM lookup where Fans.Type == "SUPPLY"
                sup_formula = (f'=INDEX({fans_c}, '
                               f'MATCH(1, INDEX(({fans_a}=A{excel_row})*({fans_b}="SUPPLY"), 0), 0))')
                sv_sys_ws.write_formula(r, 12, sup_formula, number_format)
                # N (index 13): OA CFM = F * M  (F is OA Ratio, M is Supply CFM)
                sv_sys_ws.write_formula(r, 13, f'=F{excel_row} * M{excel_row}', number_format)
                # O (index 14): Supply CFM/sf = M / D  (D is Floor Area)
                sv_sys_ws.write_formula(r, 14, f'=M{excel_row} / D{excel_row}', workbook.add_format({'num_format': '0.0#'}))
                # P (index 15): OA CFM/sf = N / D
                sv_sys_ws.write_formula(r, 15, f'=N{excel_row} / D{excel_row}', workbook.add_format({'num_format': '0.0#'}))

        # === Efficiency sheet (replicates openpyxl logic using visible TOTAL rows) ===
        if self.ss_a_data and self.ss_h_data:
            eff_ws = workbook.add_worksheet('Efficiency')
            # Header
            headers = [
                'System', 'N/A', 'Cooling Load MMBtu', 'Heating Load MMBtu',
                'Electric Heat Energy (kWh)', 'Electric Cool Energy (kWh)',
                'Heat Eff', 'Cool Eff'
            ]
            eff_ws.write_row(0, 0, headers, header_format)

            # Visible rows = those with 'TOTAL' in SS-A (like post-hide), SS-A data starts at row index 1
            ss_a_visible = [row for row in self.ss_a_data[1:] if any(isinstance(v, str) and 'TOTAL' in v for v in row)]
            ss_h_visible = [row for row in self.ss_h_data[1:] if any(isinstance(v, str) and 'TOTAL' in v for v in row)]

            max_rows = min(len(ss_a_visible), len(ss_h_visible))
            for i in range(max_rows):
                a = ss_a_visible[i]
                h = ss_h_visible[i]
                # a[0] System, a[2] Cooling Energy (MBTU), a[8] Heating Energy (MBTU)
                system = a[0]
                cool_load = a[2]
                heat_load = a[8]
                elec_heat_kwh = h[8]   # Electric Heat Energy (kWh)
                elec_cool_kwh = h[10]  # Electric Cool Energy (kWh)

                row_idx = i + 1
                eff_ws.write(row_idx, 0, system)
                eff_ws.write(row_idx, 1, "N/A")
                eff_ws.write(row_idx, 2, try_num(cool_load), number_format if is_num(cool_load) else None)
                eff_ws.write(row_idx, 3, try_num(heat_load), number_format if is_num(heat_load) else None)
                eff_ws.write(row_idx, 4, try_num(elec_heat_kwh), number_format if is_num(elec_heat_kwh) else None)
                eff_ws.write(row_idx, 5, try_num(elec_cool_kwh), number_format if is_num(elec_cool_kwh) else None)

                # Formulas
                r = row_idx + 1  # Excel 1-based
                eff_ws.write_formula(row_idx, 6, f'=IFERROR(ABS(D{r}/((E{r}*3.412)/1000)),"")', number_format)
                eff_ws.write_formula(row_idx, 7, f'=IFERROR(C{r}/((F{r}*3.412)/1000),"")', number_format)

            # Caution note (merged I1:P1)
            eff_ws.merge_range('I1:P1',
                               "⚠️ This does not account for energy consumption associated with chiller or boiler plants "
                               "(or any other water-side equipment). Therefore, exercise caution when interpreting the "
                               "efficiency values as they may not be accurate.",
                               caution_merge_fmt)

        workbook.close()


def try_convert_element_to_float(item):
    if not isinstance(item, (list, tuple)):
        li = [item]
    else:
        li = item

    out = []
    for ele in li:
        if ele is None or (isinstance(ele, str) and ele.strip() == ""):
            out.append(None)
            continue

        try:
            f = float(ele)
            if math.isnan(f) or math.isinf(f):
                out.append(None)
            else:
                out.append(f)
        except (ValueError, TypeError):
            out.append(ele)
    return out


def is_num(x):
    try:
        float(x)
        return True
    except Exception:
        return False


def try_num(x):
    try:
        return float(x)
    except Exception:
        return x


def main():
    window = tk.Tk()
    window.withdraw()
    window.iconbitmap("icon.ico")
    os.remove("icon.ico")
    filepath = filedialog.askopenfilename(
        title="Select a SIM File", filetypes=(("SIM Files", "*.SIM"), ("All Files", "*.*"))
    )

    reader = SIMFileReader(filepath)
    if reader.file_path != '':
        reader.read_file()
        reader.parse_contents()
        reader.write_excel()


if __name__ == "__main__":
    main()
