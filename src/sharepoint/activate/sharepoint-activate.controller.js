angular
    .module("Module.sharepoint.controllers")
    .controller("SharepointActivateOfficeCtrl", class SharepointActivateOfficeCtrl {

        constructor (Alerter, MicrosoftSharepointLicenseService, $stateParams, $scope, User) {
            this.alerter = Alerter;
            this.sharepointService = MicrosoftSharepointLicenseService;
            this.$stateParams = $stateParams;
            this.$scope = $scope;
            this.userService = User;
        }

        $onInit () {
            this.account = this.$scope.currentActionData;
            this.hasLicence = false;
            this.licenceOrderUrl = "https://www.ovh.com/fr/office-365/"; // default value

            this.getUser();

            this.$scope.submit = () => {
                this.$scope.resetAction();
                return this.sharepointService.updateSharepointAccount({
                    serviceName: this.$stateParams.exchangeId,
                    userPrincipalName: this.account.userPrincipalName,
                    officeLicense: true
                })
                    .then(() => {
                        this.alerter.success(
                            this.$scope.tr("sharepoint_action_activate_office_licence_success_message", this.account.userPrincipalName),
                            this.$scope.alerts.dashboard
                        );
                    })
                    .catch((err) => this.alerter.alertFromSWS(this.$scope.tr("sharepoint__action_activate_office_licence_error_message"), err, this.$scope.alerts.dashboard))
                    .finally(() => this.$scope.resetAction());
            };
        }

        getUser () {
            return this.userService.getUser()
                .then((user) => { this.licenceOrderUrl = `https://www.ovh.com/${user.ovhSubsidiary.toLowerCase()}/office-365`; })
                .catch(() => this.alerter.alertFromSWS(this.$scope.tr("sharepoint_accounts_action_sharepoint_add_error_message")));
        }
    });
