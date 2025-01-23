# Task3-Scenario2-SingleActiveCasePlugin

## Overview

This project consists of two Power Apps web resource JavaScript files and a C# .NET class library plugin. These components work together to manage primary contact information on the Case form and enforce business rules to ensure only one active case exists per customer.

## Web Resources

### 1. `primaryContactInCase.js`

**Purpose:**

-   Manages the visibility and requirement levels of the `emailaddress1`, `mobilephone`, and `primarycontactid` fields on the Case form based on the selected Customer type (Account or Contact).

**Key Functions:**

-   `hideQuickViewControlsIfDataNotPresent`: Hides or shows Quick View controls based on the presence of data.
-   `setMandatoryFields`: Sets fields as required or optional depending on whether the Customer is an Account or a Contact.
-   `hideControl`: Helper function to control the visibility of specific fields.

### 2. `autoPopulatePrimaryContact.js`

**Purpose:**

-   Automatically populates the `primarycontactid` field on the Case form when an Account is selected as the Customer by retrieving the primary contact associated with that Account.

**Key Functions:**

-   `populateContact`: Triggers on Customer field change to populate the primary contact.
-   `getPrimaryContact`: Retrieves the primary contact for a given Account using OData queries with the `Xrm.WebApi`.

## Plugin

### `Plugin.cs`

**Purpose:**

-   Ensures business rule compliance by preventing the creation of multiple active Cases for the same Customer. If an active Case already exists for a Customer, the plugin throws an exception to block the creation of a new Case.

**Key Features:**

-   Implements the `IPlugin` interface from the Microsoft.Xrm.Sdk.
-   Executes on the `Create` message for the `incident` (Case) entity.
-   Queries existing active Cases for the specified Customer using `QueryExpression`.
-   Throws an `InvalidPluginExecutionException` if an active Case is found for the Customer.

## Folder Structure

Task3-Scenario2-SingleActiveCasePlugin/
│
├── Plugin/
│ │
│ └── Plugin.cs
│
├── WebResources/
│ │
│ ├── primaryContactInCase.js
│ │
│ └── autoPopulatePrimaryContact.js
│
└── README.md

## Deployment

1. **JavaScript Web Resources:**

    - Upload `primaryContactInCase.js` and `autoPopulatePrimaryContact.js` to your Power Apps environment as web resources.
    - Add and configure these scripts on the Case form events (`OnLoad` and `OnChange` as appropriate).

2. **C# Plugin:**
    - Build the `Plugin.cs` class library.
    - Register the plugin using the Plugin Registration Tool or similar tool in your Power Apps/Dynamics 365 environment.
    - Ensure the plugin is registered on the `Create` message of the `incident` entity.

## Usage

-   **Primary Contact Management:**

    -   When a user selects a Customer on the Case form, `autoPopulatePrimaryContact.js` automatically fills in the primary contact.
    -   `primaryContactInCase.js` adjusts field requirements and visibility based on whether the Customer is an Account or a Contact.

-   **Single Active Case Enforcement:**
    -   When creating a new Case, the plugin checks for existing active Cases for the selected Customer.
    -   If an active Case exists, the creation of a new Case is blocked with an appropriate error message.

## Support

For any issues or questions, please contact the project maintainer.
