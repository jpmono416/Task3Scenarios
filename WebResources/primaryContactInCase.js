/**
 * Initialize the form by hiding controls and setting mandatory fields.
 * Should be called on form OnLoad.
 * @param {Xrm.ExecutionContext} executionContext
 */
function initialiseForm(executionContext) {
    hideQuickViewControlsIfDataNotPresent(executionContext);
    setMandatoryFields(executionContext);
}

/*
*       functions for Control visibility and mandatory fields
*/

/**
 * Hide Quick View controls if data is not present.
 * @param {Xrm.ExecutionContext} executionContext
 */
function hideQuickViewControlsIfDataNotPresent(executionContext) {
    try {
        const formContext = executionContext.getFormContext();
        const quickViewControl = getQuickViewControl(formContext, "PrimaryContactQV");

        if (!quickViewControl?.getAttribute()) {
            console.warn("Quick View Control 'PrimaryContactQV' not found or has no attributes.");
            return;
        }

        toggleControlVisibility(quickViewControl, "emailaddress1", !!getAttributeValue(quickViewControl, "emailaddress1"));
        toggleControlVisibility(quickViewControl, "mobilephone", !!getAttributeValue(quickViewControl, "mobilephone"));
    } catch (error) {
        handleError("hideQuickViewControlsIfDataNotPresent", error);
    }
}

/**
 * Set mandatory fields based on Customer type.
 * Should be called on Customer field OnChange.
 * @param {Xrm.ExecutionContext} executionContext
 */
function setMandatoryFields(executionContext) {
    try {
        const formContext = executionContext.getFormContext();
        const customerValue = getAttributeValue(formContext, "customerid");

        if (customerValue?.length === 0) {
            resetFieldRequirements(formContext);
            enableControl(formContext, "primarycontactid");
            return;
        }

        handleCustomerType(formContext, customerValue[0].entityType);
    } catch (error) {
        handleError("setMandatoryFields", error);
    }
}


/*
*       functions to populate primary contact
*/

/**
 * Populate the primary contact based on the selected Customer.
 * Should be called on Customer field OnChange.
 * @param {Xrm.ExecutionContext} executionContext
 */
async function populateContact(executionContext) {
    try {
        const formContext = executionContext.getFormContext();

        if (!isValidAccountCustomer(formContext)) {
            clearPrimaryContact(formContext);
            return;
        }
        
        // Retrieve the primary contact for the selected Account's id
        const customerValue = getAttributeValue(formContext, "customerid");
        const primaryContact = await getPrimaryContact(formatGuid(customerValue[0].id));

        if (!primaryContact) {
            clearPrimaryContact(formContext);
            return;
        }

        setPrimaryContact(formContext, primaryContact);
    } catch (error) {
        handleError("populateContact", error);
    }
}

/**
 * Retrieve the primary contact for the specified Account using OData.
 * @param {string} accountId Account ID in GUID format without braces
 * @returns {Promise<object>} Primary contact details
 */
async function getPrimaryContact(accountId) {
    try {
        // Construct the OData query to retrieve the primarycontactid with contact details
        const query = `?$select=primarycontactid&$expand=primarycontactid($select=contactid,fullname)`;
        const response = await Xrm.WebApi.retrieveRecord("account", accountId, query);

        if (response.primarycontactid) {
            return response.primarycontactid;
        }

        return null;
    } catch (error) {
        handleError("getPrimaryContact", error);
    }
}

/**
 * Set the primary contact on the form.
 * @param {Xrm.FormContext} formContext
 * @param {object} primaryContact
 */
function setPrimaryContact(formContext, primaryContact) {
    formContext.getAttribute("primarycontactid").setValue([
        {
            id: primaryContact.contactid,
            entityType: "contact",
            name: primaryContact.fullname,
        },
    ]);
}


/**
 * Clear the primarycontactid field on the form.
 * @param {Xrm.FormContext} formContext
 */
function clearPrimaryContact(formContext) {
    formContext.getAttribute("primarycontactid").setValue(null);
}

/*
*       utility functions
*/


/**
 * Handle customer type to set field requirements.
 * @param {Xrm.FormContext} formContext
 * @param {string} customerType
 */
function handleCustomerType(formContext, customerType) {
    if (customerType === "account") {
        setFieldRequirement(formContext, "primarycontactid", "required");
        enableControl(formContext, "primarycontactid", true);
    } else if (customerType === "contact") {
        setFieldRequirement(formContext, "primarycontactid", "none");
        disableControl(formContext, "primarycontactid");
    }
}

/**
 * Toggle the visibility of a control based on the provided value.
 * @param {Xrm.FormContext} formContext
 * @param {string} controlName
 * @param {boolean} isVisible
 */
function toggleControlVisibility(formContext, controlName, isVisible) {
    const control = formContext.getControl(controlName);
    if (control) {
        control.setVisible(isVisible);
    }
}

/**
 * Reset field requirements when Customer is cleared.
 * @param {Xrm.FormContext} formContext
 */
function resetFieldRequirements(formContext) {
    setFieldRequirement(formContext, "primarycontactid", "none");
    setFieldRequirement(formContext, "emailaddress1", "none");
    setFieldRequirement(formContext, "mobilephone", "none");
}

/**
 * Get a Quick View Control by name.
 * @param {Xrm.FormContext} formContext
 * @param {string} controlName
 * @returns {Xrm.Controls.QuickViewControl | null}
 */
function getQuickViewControl(formContext, controlName) {
    return formContext.ui.quickForms.get(controlName) || null;
}

/**
 * Get the value of an attribute.
 * @param {Xrm.FormContext} formContext
 * @param {string} attributeName
 * @returns {*} The value of the attribute.
 */
function getAttributeValue(formContext, attributeName) {
    const attribute = formContext.getAttribute(attributeName);
    return attribute ? attribute.getValue() : null;
}

/**
 * Set the requirement level of a field.
 * @param {Xrm.FormContext} formContext
 * @param {string} attributeName
 * @param {string} requirementLevel ('none', 'required', 'recommended')
 */
function setFieldRequirement(formContext, attributeName, requirementLevel) {
    const attribute = getAttributeValue(formContext, attributeName);
    if (attribute) {
        attribute.setRequiredLevel(requirementLevel);
    }
}

/**
 * Disable a control.
 * @param {Xrm.FormContext} formContext
 * @param {string} controlName
 */
function disableControl(formContext, controlName) {
    const control = formContext.getControl(controlName);
    if (control) {
        control.setDisabled(true);
    }
}

/**
 * Enable a control.
 * @param {Xrm.FormContext} formContext
 * @param {string} controlName
 * @param {boolean} [enable=true]
 */
function enableControl(formContext, controlName, enable = true) {
    const control = formContext.getControl(controlName);
    if (control) {
        control.setDisabled(!enable);
    }
}

/**
 * Format GUID by removing braces.
 * @param {string} guid GUID with or without braces
 * @returns {string} GUID without braces
 */
function formatGuid(guid) {
    return guid.replace("{", "").replace("}", "");
}


/**
 * Check if the customer is valid and an account.
 * @param {Xrm.FormContext} formContext
 * @returns {boolean} True if the customer is valid and an account, false otherwise.
 */
function isValidAccountCustomer(formContext) {
    const customerValue = getAttributeValue(formContext, "customerid");
    return customerValue && customerValue.length > 0 && customerValue[0].entityType === "account";
}

/**
 * Handle errors by logging and displaying an alert.
 * @param {string} functionName
 * @param {Error} error
 */
function handleError(functionName, error) {
    console.error(`Error in ${functionName}:`, error);
    Xrm.Navigation.openAlertDialog({ text: `An unexpected error occurred: ${error.message}` });
}