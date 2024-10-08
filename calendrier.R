library(readxl)
library(dplyr)
library(lubridate)
library(tidyr)
library(stringr)

{
  tmp <- tempfile()
  download.file("https://nextcloud.centralesupelec.fr/s/y2oSoJQTGQ4D9RR/download/SETI-public.xlsx",
                tmp,
                method = "wget",
                quiet = TRUE)
  tmp
} %>%
  read_xlsx(range = cell_limits(c(2L, 1L),
                                c(NA_integer_, 11L),
                                "Edt 2024-25"),
            col_types = c("date", rep("text", 10L))) %>%
  rename(date = ...1,
         Lundi_bis = ...3,
         Mardi_bis = ...5,
         Mercredi_bis = ...7,
         Jeudi_bis = ...9,
         Vendredi_bis = ...11) %>%
  suppressWarnings() %>%
  fill(date, .direction = "down") %>%
  mutate(date = floor_date(as_datetime(round_date(date, unit = "day"), tz = "Europe/Paris"), unit = "day")) %>%
  group_by(ymd(date)) %>%
  mutate(matin = 1:n() <= 4L) %>%
  ungroup() %>%
  select(date,
         Lundi, Mardi, Mercredi, Jeudi, Vendredi,
         Lundi_bis, Mardi_bis, Mercredi_bis, Jeudi_bis, Vendredi_bis,
         matin) %>%
  pivot_longer(cols = c(Lundi, Mardi, Mercredi, Jeudi, Vendredi,
                        Lundi_bis, Mardi_bis, Mercredi_bis, Jeudi_bis, Vendredi_bis),
               names_to = "Jour",
               values_to = "summary") %>%
  filter(! is.na(summary) & summary != "") %>%
  mutate(date = case_when(
    Jour == "Lundi" ~ date,
    Jour == "Mardi" ~ date + days(1L),
    Jour == "Mercredi" ~ date + days(2L),
    Jour == "Jeudi" ~ date + days(3L),
    Jour == "Vendredi" ~ date + days(4L),
    Jour == "Lundi_bis" ~ date,
    Jour == "Mardi_bis" ~ date + days(1L),
    Jour == "Mercredi_bis" ~ date + days(2L),
    Jour == "Jeudi_bis" ~ date + days(3L),
    Jour == "Vendredi_bis" ~ date + days(4L)
  )) %>%
  mutate(match = str_match(summary, "([0-9]{1,2})\\s?h\\s?([0-9]{0,2})\\s?[\\s/|\\-:]\\s?([0-9]{1,2})\\s?h\\s?([0-9]{0,2})"),
         heure_debut = hours(as.integer(match[,2L])) + minutes(replace_na(as.integer(match[,3L]), 0L)),
         heure_fin = hours(as.integer(match[,4L])) + minutes(replace_na(as.integer(match[,5L]), 0L))) %>%
  mutate(matin_impute = (is.na(heure_debut) | is.na(heure_fin)) & matin,
         aprem_impute = (is.na(heure_debut) | is.na(heure_fin)) & ! matin) %>%
  mutate(summary = case_when(matin_impute ~ str_c("[Heure imputée matin] ", summary),
                             aprem_impute ~ str_c("[Heure imputée aprem] ", summary),
                             .default = summary),
         heure_debut = case_when(matin_impute ~ hours(8L) + minutes(45L),
                                 aprem_impute ~ hours(13L) + minutes(45L),
                                 .default = heure_debut),
         heure_fin = case_when(matin_impute ~ hours(12L) + minutes(30L),
                               aprem_impute ~ hours(17L) + minutes(45L),
                               .default = heure_fin)) %>%
  mutate(summary = str_replace_all(summary, "(\r)?(\n)", " "),
         debut = format(with_tz(date + heure_debut, "UTC"), "%Y%m%dT%H%M%SZ"),
         fin =  format(with_tz(date + heure_fin, "UTC"), "%Y%m%dT%H%M%SZ"),
         stamp = format(with_tz(now(), "UTC"), "%Y%m%dT%H%M%SZ")) %>%
  mutate(Event = str_c(
    "BEGIN:VEVENT",
    str_c("UID:", 1L:n(), "@uid.com"),
    str_c("DTSTART:", debut),
    str_c("DTEND:", fin),
    str_c("DTSTAMP:", stamp),
    str_c("SUMMARY:", summary),
    "END:VEVENT",
    sep = "\r\n")) %>%
  pull(Event) %>%
  str_c(collapse = "\r\n") %>%
  {
    str_c(
      "BEGIN:VCALENDAR",
      "VERSION:2.0",
      "PRODID:-//UPS//Seti//FR",
      "BEGIN:VTIMEZONE",
      "TZID:UTC",
      "BEGIN:STANDARD",
      "DTSTART:20240101T000000",
      "TZOFFSETFROM:+0000",
      "TZOFFSETTO:+0000",
      "TZNAME:UTC",
      "END:STANDARD",
      "END:VTIMEZONE",
      .,
      "END:VCALENDAR",
      sep = "\r\n")
  } %>%
  writeLines("edt.ics")
