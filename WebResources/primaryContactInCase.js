/**
 * Initialize the form by hiding controls and setting mandatory fields.
 * Should be called on form OnLoad.
 * @param {Xrm.ExecutionContext} executionContext 
 */
function initialiseForm(executionContext) {
    hideQuickViewControlsIfDataNotPresent(executionContext);
    setMandatoryFields(executionContext);
}

/**
 * Hide Quick View controls if data is not present.
 * @param {Xrm.ExecutionContext} executionContext 
 */
function hideQuickViewControlsIfDataNotPresent(executionContext) {
    try {
        const formContext = executionContext.getFormContext();
        const quickViewControl = getQuickViewControl(formContext, "PrimaryContactQV");

        if (!quickViewControl || !quickViewControl.getAttribute()) {
            console.warn("Quick View Control 'PrimaryContactQV' not found or has no attributes.");
            return;
        }

        // Check if 'emailaddress1' has data
        const emailValue = getAttributeValue(formContext, "emailaddress1");
        toggleControlVisibility(formContext, "emailaddress1", !!emailValue);

        // Check if 'mobilephone' has data
        const phoneValue = getAttributeValue(formContext, "mobilephone");
        toggleControlVisibility(formContext, "mobilephone", !!phoneValue);
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

        if (!customerValue || customerValue.length === 0) {
            resetFieldRequirements(formContext);
            enableControl(formContext, "primarycontactid");
            return;
        }

        const customerType = customerValue[0].entityType;

        if (customerType === "account") {
            setFieldRequirement(formContext, "primarycontactid", "required");
            enableControl(formContext, "primarycontactid", true);
        } else if (customerType === "contact") {
            setFieldRequirement(formContext, "primarycontactid", "none");
            disableControl(formContext, "primarycontactid");
        }
    } catch (error) {
        handleError("setMandatoryFields", error);
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
    const attribute = formContext.getAttribute(attributeName);
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
 * Handle errors by logging and displaying an alert.
 * @param {string} functionName 
 * @param {Error} error 
 */
function handleError(functionName, error) {
    console.error(`Error in ${functionName}:`, error);
    Xrm.Navigation.openAlertDialog({ text: `An unexpected error occurred: ${error.message}` });
}