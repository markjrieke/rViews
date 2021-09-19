README
================
Mark Rieke
9/19/2021

Using R for business often involves generating recurring reports that
are emailed to stakeholders. While R scripts take the legwork out of
creating the report, there is still quite a bit of labor involved in
rerunning the script and sending out an email with the report attached.
Third party automations, like GitHub Actions LINK, can run the R script
regularly, but companies working with sensitive data may require that
everything stays within the existing infrastructure. Many companies use
Outlook as the email platform and although there is an R package for
sending email, mailR LINK, companies with software security that errs on
the side of caution may not work well with the package. Analysts,
therefore, may find themselves in a unique scenario requiring localized
automation of report building and emailing via Outlook. The solution?
Convincing R, VBA, and Window’s Task Scheduler to work together to meet
your automation needs.

In this article, we’ll look at three basic workflows: \* Generating a
report with R \* Automatically running a script with the Task Scheduler
\* Automatically sending emails via Outlook with VBA

This article is heavily informed by PERSON LINK and PERSON LINK articles
on the Task Scheduler and VBA in Outlook, with some stumbling blocks
pointed out along the way.
