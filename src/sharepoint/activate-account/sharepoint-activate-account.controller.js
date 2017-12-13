angular
    .module("Module.sharepoint.controllers")
    .controller("SharepointActivateAccountCtrl", class SharepointActivateAccountCtrl {

        constructor (Alerter, MicrosoftSharepointLicenseService, $stateParams, $scope) {
            this.alerter = Alerter;
            this.sharepointService = MicrosoftSharepointLicenseService;
            this.$stateParams = $stateParams;
            this.$scope = $scope;
        }

        $onInit () {
            this.account = this.$scope.currentActionData;

            this.$scope.submit = () => {
                this.$scope.resetAction();
                return this.sharepointService.activateSharepointOnAccount(this.$stateParams.exchangeId, this.account.userPrincipalName)
                    .then(() => {
                        this.alerter.success(
                            this.$scope.tr("sharepoint_action_activate_account_success_message", this.account.userPrincipalName),
                            this.$scope.alerts.main
                        );
                    })
                    .catch((err) => {
                        this.alerter.alertFromSWS(
                            this.$scope.tr("sharepoint_action_activate_account_error_message", this.account.userPrincipalName),
                            err,
                            this.$scope.alerts.main
                        );
                    })
                    .finally(() => this.$scope.resetAction());
            };
        }

    });
