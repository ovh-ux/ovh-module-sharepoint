angular
    .module("Module.sharepoint.controllers")
    .controller("SharepointUpdatePasswordCtrl", class SharepointUpdatePasswordCtrl {
        constructor (Alerter, ExchangePassword, MicrosoftSharepointLicenseService, $stateParams, $scope) {
            this.alerter = Alerter;
            this.exchangePassword = ExchangePassword;
            this.microsoftSharepointLicense = MicrosoftSharepointLicenseService;
            this.$stateParams = $stateParams;
            this.$scope = $scope;
        }

        $onInit () {
            this.account = this.$scope.currentActionData;
            this.exchangeId = this.$stateParams.exchangeId;
            this.passwordTooltip = null;

            this.retrievingMSService();
            this.retrievingExchangeOrganization();

            this.$scope.updatingSharepointPassword = () => this.updatingSharepointPassword();
        }

        updatingSharepointPassword () {
            const model = {
                serviceName: this.exchangeId,
                userPrincipalName: this.account.userPrincipalName,
                password: this.account.password
            };

            return this.microsoftSharepointLicense
                .updatingSharepointPasswordAccount(model)
                .then(() => {
                    this.alerter.success(this.$scope.tr("sharepoint_ACTION_update_password_confirm_message_text", this.account.userPrincipalName), this.$scope.alerts.dashboard);
                })
                .catch((err) => {
                    this.alerter.alertFromSWS(this.$scope.tr("sharepoint_ACTION_update_password_error_message_text"), err, this.$scope.alerts.dashboard);
                })
                .finally(() => {
                    this.$scope.resetAction();
                });
        }

        retrievingMSService () {
            return this.microsoftSharepointLicense
                .retrievingMSService(this.exchangeId)
                .then((exchange) => {
                    this.exchange = exchange;
                    this.setPasswordTooltipMessage();
                });
        }

        setPasswordTooltipMessage () {
            const messageId = this.exchange.complexityEnabled ? "sharepoint_ACTION_update_password_complexity_message_all" : "sharepoint_ACTION_update_password_complexity_message_length";
            this.passwordTooltip = this.$scope.tr(messageId, [this.exchange.minPasswordLength]);
        }

        retrievingExchangeOrganization () {
            return this.microsoftSharepointLicense
                .retrievingExchangeOrganization(this.exchangeId)
                .then((organization) => {
                    this.hasAssociatedExchange = !_.isEmpty(organization);
                });
        }

        setPasswordsFlag (selectedAccount) {
            this.differentPasswordFlag = false;
            this.simplePasswordFlag = false;
            this.containsNameFlag = false;
            this.containsSameAccountNameFlag = false;

            selectedAccount.password = selectedAccount.password || "";
            selectedAccount.passwordConfirmation = selectedAccount.passwordConfirmation || "";

            if (selectedAccount.password !== selectedAccount.passwordConfirmation) {
                this.differentPasswordFlag = true;
            }

            if (selectedAccount.password.length > 0) {
                this.simplePasswordFlag = !this.exchangePassword.passwordSimpleCheck(selectedAccount.password, true, this.exchange.minPasswordLength);

                // see the password complexity requirements of Windows Server (like Exchange)
                // https://technet.microsoft.com/en-us/library/hh994562%28v=ws.10%29.aspx
                if (this.exchange.complexityEnabled) {
                    this.simplePasswordFlag = this.simplePasswordFlag || !this.exchangePassword.passwordComplexityCheck(selectedAccount.password);

                    if (selectedAccount.displayName) {
                        this.containsNameFlag = this.exchangePassword.passwordContainsName(
                            selectedAccount.password,
                            selectedAccount.displayName
                        );
                    }

                    if (!this.containsNameFlag && selectedAccount.login) {
                        if (_.some(selectedAccount.password, selectedAccount.login)) {
                            this.containsNameFlag = true;
                        }
                    }

                    if (selectedAccount.samaccountName && _.some(selectedAccount.password, selectedAccount.samaccountName)) {
                        if (!this.containsSamAccountNameLabel) {
                            this.containsSamAccountNameLabel = this.$scope.tr("exchange_ACTION_update_account_step1_password_contains_samaccount_name", [selectedAccount.samaccountName]);
                        }

                        this.containsSamAccountNameFlag = true;
                    } else {
                        this.containsSamAccountNameFlag = false;
                    }
                }
            }
        }
    });
