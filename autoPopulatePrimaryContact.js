/**
 * Populate the primary contact based on the selected Customer.
 * @param {Xrm.ExecutionContext} executionContext Form execution context
 */
function populateContact(executionContext) {
    try {
        const formContext = executionContext.getFormContext();
        const customerValue = formContext.getAttribute("customerid").getValue();

        if (!customerValue || customerValue.length === 0) {
            return;
        }

        const customerType = customerValue[0].entityType;

        // If Account → populate primarycontactid
        if (customerType === "account") {
            const accountId = customerValue[0].id;
            getPrimaryContact(accountId)
                .then((primaryContact) => {
                    if (primaryContact) {
                        formContext.getAttribute("primarycontactid").setValue([
                            {
                                id: primaryContact.contactid,
                                entityType: "contact",
                                name: primaryContact.fullname,
                            },
                        ]);
                    }
                })
                .catch((error) => {
                    console.error("Error retrieving primary contact: ", error);
                    Xrm.Navigation.openAlertDialog({
                        text:
                            "An error occurred while retrieving the primary contact. " +
                            error.message,
                    });
                });
        }
    } catch (error) {
        console.error("Error in populateContact: ", error);
        Xrm.Navigation.openAlertDialog({ text: "An unexpected error occurred. " + error.message });
    }
}

/**
 * Retrieve the primary contact for the specified Account using OData.
 * @param {string} accountId Account ID
 * @returns {Promise<object>} Primary contact details
 */
function getPrimaryContact(accountId) {
    // Ensure the accountId is in GUID format without braces
    //const formattedAccountId = accountId.replace("{", "").replace("}", "");

    // Construct the OData query
    const query = `?$select=contactid,fullname&$filter=accountid eq ${accountId} and isprimary eq true&$top=1`;

    return Xrm.WebApi.retrieveMultipleRecords("contact", query).then((response) => {
        if (response.entities && response.entities.length > 0) {
            return response.entities[0];
        }
        return null;
    });
}
