import collections

class DipendentiCloud:
    def __init__(self, path):
        self.path = path

    def parse_file(self):
        return None

class Dipendente:
    def __init__(self, name, matricola, movimenti):
        self.name = name
        self.matricola = matricola
        self.movimenti = movimenti
        self.codice_ufficiale_dipendente = self.compute_codice()
        self.codice_azienda_ufficiale = "000053"
        self.calendar = self.create_calendar()

    def compute_codice(self):
        return format(self.matricola, "07")

    def create_calendar(self):
        calendario = {}

        for movimento in self.movimenti:
            date = movimento.date
            nr_ore = movimento.nr_ore
            giust = movimento.giustificativo
            # if giust == "Presenza ordinaria prevista":
            #     continue
            # else:
            if nr_ore != 0 and str(nr_ore) != 'nan':
                if date in calendario:
                    simple_dict = calendario[date]
                    simple_dict[giust] = nr_ore
                else:
                    simple_dict = {giust: nr_ore}
                    calendario[date] = simple_dict
            else:
                continue
        return collections.OrderedDict(sorted(calendario.items()))
        # from keys remove also the duplicates in "days" list
        # new_dict = dict.fromkeys(days, days_dict)
        # return new_dict
# days_dict[date] = {giust_1: nr_ore, giust_2: nr_ore}

class Movimento:
    def __init__(self, giustificativo, date, nr_ore):
        self.giustificativo = giustificativo
        self.date = date
        self.nr_ore = nr_ore

