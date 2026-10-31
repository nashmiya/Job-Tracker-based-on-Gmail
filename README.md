# Sample outputs

Fictional data showing what the tracker produces. Every company, person, address and ID here is invented
(the company names are standard placeholder names such as Contoso, Northwind and Fabrikam; email addresses use the reserved `.example` domain).

| File | Matches tab | Columns |
|---|---|---|
| `sample_applications.csv` | Applications | Key, Company, Job title, Location, Country, Channel, Applied on, Last update, Status (auto), Status override, Final status, Days since update, Emails, Gmail link, Notes |
| `sample_events.csv` | Events | Message ID, Thread ID, Date, From, Subject, Type, Company, Job title, Location, Country, Channel, App key, Gmail link |
| `sample_summary.csv` | Summary | Metric, All, Germany, Other countries; conversion by year |

Things to notice:

- One application has a manual **Status override** (`Interviewing`) and a note, showing how to record a phone screen that produced no email.
- `Final status` turns `Applied` into `No reply` after 42 days without any email.
- An application that had an interview and then a rejection appears as `Rejected after interview`.

You can import these CSVs into a blank Google Sheet (**File → Import**) to explore the layout before connecting your own Gmail.
