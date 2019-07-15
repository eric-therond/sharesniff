# Sharesniff

Shareniff is a tool that can be used to:
- detect sensitive data in documents
- to approximate the value of a document
- forensic

It analyzes documents to find strings, called * artifacts *, representative of sensitive data. It is complementary to a tool like grep because it understands among others Office documents (Word, Excel, Outlook, Powerpoint ...).

## Configuration

A config.json file at the root can be set:
- *report_name*: the name of your analysis
- *max_urls*: the max number of urls analyzed when the tool is in web crawling mode
- *strings_artifacts*: specifies the list of artifacts
	- name: the name of the artifact
	- string: the string of characters to recognize (the use of regex is possible)
	- dvs: the digital value score, the sensitivity of the data represented by the string

## Use

```bash
.\sharesniff.ps1 typecrawler path
```

Where typecrawler can be one of the modes of analysis below and path your analysis target.

## Crawlers
### Filesystem

The analysis of a file or file:

```bash
.\sharesniff.ps1 filesystem c:\folder
```

### Sharepoint

The analysis of a Sharepoint site (requires to run sharesniff on Sharepoint Server):

```bash
.\sharesniff.ps1 sharepoint url
```

### Website

Analysis of a website:

```bash
.\sharesniff.ps1 website url
```

## Résultats

The result of the analysis is written in the ./report/datatable.js file and a web HMI is available here ./report/template.html
