Automated Email Reports with R, VBA, and the Task Scheduler
================
Mark Rieke
9/19/2021

Using R for business often involves generating recurring reports that
are emailed to stakeholders. While R scripts take the legwork out of
creating the report, there is still quite a bit of labor involved in
rerunning the script and sending out an email with the report attached.
Third party automation services, like [GitHub
Actions](https://github.com/features/actions), can run the R script
regularly, but companies working with sensitive data may require that
everything stays within the existing infrastructure. Many companies use
Outlook as the email platform and although there is an R package for
sending email, [mailR](https://rpremraj.github.io/mailR/), companies
with software security that errs on the side of caution may not work
well with the package. Analysts, therefore, may find themselves in a
unique scenario requiring localized automation of report building and
emailing via Outlook. The solution? Convincing R, VBA, and Window’s Task
Scheduler to work together to meet your automation needs.

In this article, we’ll look at three basic workflows:

-   Generating a report with R
-   Automatically running a script with the Task Scheduler
-   Automatically sending emails via Outlook with VBA

> This article is heavily informed by [Sean Carney’s article on running
> scripts with the Task
> Scheduler](http://www.seancarney.ca/2020/10/11/scheduling-r-scripts-to-run-automatically-in-windows/)
> and [Shirly Zhang’s article on scheduling emails with
> VBA](https://www.datanumen.com/blogs/auto-send-recurring-email-periodically-outlook-vba/).
> This article merges the two concepts and points out additional
> stumbling blocks, but the original articles are well worth viewing.

## Generating a Report with R

Suppose we want to create a .csv file containing a weekly summary of the
daily new COVID cases in Texas. Our script to generate the report may
look something like this:

``` r
library(dplyr)
library(readr)

# set working directory
setwd("C:/path/to/your/directory")

# read in data
nyt_covid <- read_csv("https://raw.githubusercontent.com/nytimes/covid-19-data/master/us-states.csv")

# get texas' daily new cases in the past week
nyt_covid <- 
  nyt_covid %>%
  group_by(state) %>%
  mutate(new_cases = cases - lag(cases),
         new_deaths = deaths - lag(deaths)) %>%
  filter(date >= Sys.Date() - 7,
         state == "Texas")

# save file
nyt_covid %>%
  write_csv("new_cases_tx.csv")
```
