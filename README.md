Habitat Evidence Retrieval for Metabarcoding Taxa

Overview

This repository contains scripts for retrieving habitat and environment related information from external biodiversity resources. The extracted evidence supports the manual assignment of habitat categories to taxa detected through metabarcoding. The scripts collect and organise supporting records from websites; they do not determine the final habitat classifications automatically.

Data Sources

Sources were selected for different taxonomic groups based on the availability of relatively structured habitat or environmental information:

WoRMS (World Register of Marine Species)

IUCN Red List

GBIF (Global Biodiversity Information Facility)

NatureServe Explorer

FungalTraits

Burke Herbarium

Go Botany

The Workflow

Evidence extraction: Where available, the scripts retrieve species names(sometimes genus if available), habitat or environment descriptions, and the corresponding source URLs.

Local caching: Up to 10 species records per genus are retained in local cache files, with their extracted information, to limit the amount of data processed and preserve the retrieved evidence for review.

Evidence selection: Because the full cached information may be too extensive for an Excel cell, up to 3 species records per genus are randomly selected(random seed 42) for inclusion in the spreadsheet's Evidence column. Each entry identifies the species and its reported habitat or environment, together with the source and URL. For example:

[Raphidiophrys elegans] /Environment: freshwater /source URL/

Manual classification: The Evidence column is reviewed manually(the website info can be wrong sometime), and a final habitat category is assigned based on the available evidence. Taxa for which the scripts cannot retrieve sufficient supporting information are flagged for manual checking.

Spreadsheet output: The evidence, source references, and final habitat assignments are recorded in an .xlsx file and subjected to manual checking and correction.
