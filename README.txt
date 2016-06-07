# VDI-Auto-Provision

This is a set of scripts I special designed for my previous company, confidential information removed (as I know).
**NOTE**: The project is quite ancient, based on old releases and specially designed for the company, I really dont think it can be directly used for whom download it. Most valuable is reference only.

Environments:
	Windows 2008 R2
	XenDesktop 5.6
	SharePoint 2010
	Powershell 2.0
	NetApp as NAS (version forgot)

Schema:
	frontend - SharePoint 2010 - provide web UI and export VDI creation requests to CSV file
	backend - XenDesktop 5.6 - read VDI creation request CSV file and process creating VDI

Sharepoint references:
	# The sharepoint BAK file most likely contains confidential information so I can't upload yet
	# Powershell export shartpoint page contents to CSV file
	# Powershell import VDI reports

XenDesktop references:
	# Invoke web service to do rapid clone for VDI master template.
	# Reclaim NetApp space
	# Read VDI creation request CSV file from sharepoint server and provision VDIs

- Larry.Song@outlook.com
