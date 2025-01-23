/**
 * Hide Quick View controls if data is not present
 * @param {*} executionContext Form execution context
 */
function hideQuickViewControlsIfDataNotPresent(executionContext) {
    try {
        const formContext = executionContext.getFormContext();

        // Retrieve the related Account’s primary contact via Quick View
        const quickViewControl = formContext.ui.quickForms.get("PrimaryContactQV");
        if (!quickViewControl?.getAttribute()) {
            return;
        }

        // Conditionally render the controls
        this.hideControl(quickViewControl, "emailaddress1");
        this.hideControl(quickViewControl, "mobilephone");
    } catch (error) {
        console.error("Error in hideQuickViewControlsIfDataNotPresent: ", error);
        Xrm.Navigation.openAlertDialog({ text: "An unexpected error occurred. " + error.message });
    }
}

/**
 * Set mandatory fields based on Customer type
 * @param {*} executionContext Form execution context
 */
function setMandatoryFields(executionContext) {
    const formContext = executionContext.getFormContext();
    const customerValue = formContext.getAttribute("customerid").getValue();

    if (!customerValue || customerValue.length === 0) {
        return;
    }

    const customerType = customerValue[0].entityType;

    // If Account → make primarycontactid required
    formContext
        .getAttribute("primarycontactid")
        .setRequiredLevel(customerType === "account" ? "required" : "none");
    formContext.getControl("primarycontactid").setDisabled(customerType === "contact");
}

/**
 * Hide the control if the value is not present
 * @param {*} parentFormContext Quick View control
 * @param {string} controlName Name of the control to hide
 */
this.hideControl = function (parentFormContext, controlName) {
    var control = parentFormContext.getControl(controlName);
    if (control) {
        if (!parentFormContext.getAttribute(controlName).getValue()) {
            control.setVisible(false);
        } else {
            control.setVisible(true);
        }
    }
};