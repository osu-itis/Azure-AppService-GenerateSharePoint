---
Graph:
  Endpoint: https://graph.microsoft.com/
  Versions: v1.0
  Rest Options: Post, Get
PowerShell Compatibility:
  PS 5.1: ✔
  PS 7.0+: ✔
---

# Generate SharePoint

This repository contains the code used to make a new SharePoint site using a single API call. This was created for use with TDx web request features to automate workflows.

- [Generate SharePoint](#generate-sharepoint)
  - [Setup](#setup)
    - [Azure Setup](#azure-setup)
    - [TDx Setup](#tdx-setup)
  - [Workflow and Usage](#workflow-and-usage)
  - [Rest Examples](#rest-examples)
    - [Example Hostname](#example-hostname)

## Setup

### Azure Setup

1. Create a new Azure App Registration.
   1. Make note of the application (Client ID), we will need this value later.
   2. Make note of the Directory (tenant) ID, we will need this value later.
   3. Add a new client secret and make note of the value, we will need this value later.
   4. Add the needed graph permissions and grant admin consent:
      | Permissions name | Type | Description | Admin Consent Required |
      | --- | --- | --- | --- |
      | Group.Create | Application | Create groups | Yes |
      | Group.Read.All | Application | Read all groups | Yes |
      | Group.ReadWrite.All | Application | Read and write all groups | Yes |
      | GroupMember.Read.All | Application | Read all group memberships | Yes |
      | GroupMember.ReadWrite.All | Application | Read and write all group memberships | Yes |
      | Sites.Manage.All | Application | Create, edit, delete items and lists in all site collections | Yes |
      | Sites.Read.All | Application | Read items in all site collections | Yes |
      | Sites.ReadWrite.All | Application | Read and write items in all site collections | Yes |
      | User.Read.All | Application | Read all users' full profiles | Yes |
2. Create a new Azure Resource group.
3. Create a new function app.
   1. Add the following application settings:
      1. ClientID.
      2. ClientSecret.
      3. TenantID.
   2. Configure TLS/SSL settings:
      1. Enable HTTPS Only.
   3. Configure Deployment Center:
      1. Set code source to Manual deployment - External Git
         1. Specify the repository and branch and wait for sync to complete.
   4. Review the functions section and confirm that the expected functions are listed.
      1. Review each function and make note of the function URL value, we will need this value later.

### TDx Setup

1. Create a form.
   1. Add two new fields, one for the sharepoint site name, one for the sharepoint site description.
2. Create a new Web Service Method.
   1. Set parameters:
      | Name | Data Type | Source | Source Property |
      | --- | --- | --- | --- |
      | owner | string | from ticket | requestor email |
      | ticket ID | string | from ticket | id |
      | displayName | string | from ticket | sharepoint - name |
      | description | string | from ticket | sharepoint - description |
   2. Set the body block:
      ```BodyBlock
      {
          "ticketID":"{{ticketID}}",
          "owner":"{{owner}}",
          "displayName":"{{displayName}}",
          "description":"{{description}}"
      }
      ```
3. Create a service.
4. Create and build a workflow that uses the web service method.

## Workflow and Usage

- `NewSharePointSite` Function is triggered via a Post request.

## Rest Examples

View the readme file within each function's respective folder for more information.

### Example Hostname

```text
The name is based off of whatever the Azure Function App Service name is:
  Host: https://<FUNCTIONAPPNAME>.azurewebsites.net

Port 7071 is currently the default port when using the local azure function apps for testing:
  Host: localhost:7071

```
