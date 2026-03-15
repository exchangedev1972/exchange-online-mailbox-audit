# Exchange Online Mailbox Audit

PowerShell script for exporting Exchange Online mailbox audit details to CSV.

## Overview

This script retrieves mailbox configuration and permission details from Exchange Online to support administrative review, governance reporting, and operational analysis.

It collects mailbox size, alias addresses, delegated permissions, send-as permissions, tenant routing addresses, and forwarding configuration, then exports the results to a structured CSV report.

## Features

- Enumerates all Exchange Online mailboxes
- Excludes specified mailbox addresses
- Reports mailbox size statistics
- Captures alias addresses (excluding system and routing addresses)
- Identifies tenant routing (mail.onmicrosoft.com) addresses
- Reports Full Access permissions
- Reports Send As permissions
- Reports mailbox forwarding configuration
- Exports results to CSV for reporting and analysis

## Requirements

- PowerShell 7 recommended
- ExchangeOnlineManagement module
- Exchange Online administrative permissions

## Install Module

Install the Exchange Online PowerShell module:

Install-Module ExchangeOnlineManagement -Scope CurrentUser

## Usage

Connect to Exchange Online:

Connect-ExchangeOnline

Run the script:

.\Export-AllMailboxesReport.ps1

## Output

The script generates the following file in the working directory:

AllMailboxesReport.csv

The exported report includes:

- Display Name
- Primary SMTP Address
- Alias Addresses
- Tenant Routing Address
- Mailbox Size
- Full Access Users
- Send As Users
- Forwarding Address
- Forwarding SMTP Address
- DeliverToMailboxAndForward setting

## Purpose

This script is intended to support:

- Exchange Online administrative reporting
- Permission auditing
- Mail flow and forwarding review
- Tenant routing verification
- Operational mailbox governance

## Notes

- Script is designed for reporting and audit scenarios
- System-generated proxy addresses are filtered from alias output
- Requires appropriate Exchange Online administrative permissions
