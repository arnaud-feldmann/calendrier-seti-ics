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
  mutate(Date = round_date(as_datetime(Date, tz = "Europe/Paris"), unit = "hour")) %>%
  group_by(Date) %>%
  summarise(across(everything(),
                   ~str_c(.x[!is.na(.x)], collapse = " || "))) %>%
  ungroup() %>%
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
  filter(! is.na(Edt) & Edt != "") %>%
  mutate(Edt = str_replace_all(Edt, "(\r)?(\n)", " "),
         debut = format(with_tz(Date + hours(6L) + minutes(45L), "UTC"), "%Y%m%dT%H%M%SZ"),
         fin =  format(with_tz(Date + hours(15L) + minutes(30L), "UTC"), "%Y%m%dT%H%M%SZ"),
         stamp = str_c("UID:", 1L:n(), "@uid.com")) %>%
  mutate(Event = str_c(
    "BEGIN:VEVENT",
    str_c("UID:", 1L:n(), "@uid.com"),
    str_c("DTSTART;TZID=UTC:", debut),
    str_c("DTEND;TZID=UTC:", fin),
    str_c("DTSTAMP;TZID=UTC:", stamp),
    str_c("SUMMARY:", str_sub(Edt, 1L, 30L)),
    str_c("DESCRIPTION:", Edt),
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
