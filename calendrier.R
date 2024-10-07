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
  suppressWarnings() %>%
  select(Date = ...1, Lundi, Mardi, Mercredi, Jeudi, Vendredi) %>%
  filter(! (is.na(Lundi) & is.na(Mardi) & is.na(Mercredi) &
              is.na(Jeudi) & is.na(Vendredi))) %>%
  fill(Date, .direction = "down") %>%
  mutate(Date = floor_date(as_datetime(round_date(Date, unit = "day"), tz = "Europe/Paris"), unit = "day")) %>%
  pivot_longer(cols = c(Lundi, Mardi, Mercredi, Jeudi, Vendredi),
               names_to = "Jour",
               values_to = "Edt") %>%
  mutate(Date = case_when(
    Jour == "Lundi" ~ Date,
    Jour == "Mardi" ~ Date + days(1L),
    Jour == "Mercredi" ~ Date + days(2L),
    Jour == "Jeudi" ~ Date + days(3L),
    Jour == "Vendredi" ~ Date + days(4L)
  )) %>%
  select(! Jour) %>%
  mutate(match = str_match(Edt, "([0-9]{1,2})h([0-9]{0,2})\\s?[\\s/|\\-:]\\s?([0-9]{1,2})h([0-9]{0,2})"),
         heure_debut = hours(as.integer(match[,2L])) + minutes(replace_na(as.integer(match[,3L]), 0L)),
         heure_fin = hours(as.integer(match[,4L])) + minutes(replace_na(as.integer(match[,5L]), 0L))) %>%
  mutate(Edt = if_else(is.na(heure_debut) | is.na(heure_fin), str_c("[PAS D'HEURE] ", Edt), Edt)) %>%
  mutate(heure_debut = replace_na(heure_debut, hours(8L) + minutes(45L)),
         heure_fin = replace_na(heure_fin, hours(17L) + minutes(30L))) %>%
  filter(! is.na(Edt) & Edt != "") %>%
  mutate(Edt = str_replace_all(Edt, "(\r)?(\n)", " "),
         debut = format(with_tz(Date + heure_debut, "UTC"), "%Y%m%dT%H%M%SZ"),
         fin =  format(with_tz(Date + heure_fin, "UTC"), "%Y%m%dT%H%M%SZ"),
         stamp = format(with_tz(now(), "UTC"), "%Y%m%dT%H%M%SZ")) %>%
  mutate(Event = str_c(
    "BEGIN:VEVENT",
    str_c("UID:", 1L:n(), "@uid.com"),
    str_c("DTSTART:", debut),
    str_c("DTEND:", fin),
    str_c("DTSTAMP:", stamp),
    str_c("SUMMARY:", Edt),
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
