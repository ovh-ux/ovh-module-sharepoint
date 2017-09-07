angular
    .module("Module.sharepoint.controllers")
    .controller("SharepointDeleteAccountCtrl", class SharepointDeleteAccountCtrl {

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
                this.sharepointService.deleteSharepointAccount(this.$stateParams.exchangeId, this.account.userPrincipalName)
                    .then(() => {
                        this.alerter.success(
                            this.$scope.tr("sharepoint_account_action_sharepoint_remove_success_message", this.account.userPrincipalName),
                            this.$scope.alerts.dashboard
                        );
                    })
                    .catch((err) => {
                        this.alerter.alertFromSWS(
                            this.$scope.tr("sharepoint_account_action_sharepoint_remove_error_message"), err, this.$scope.alerts.dashboard
                        );
                    })
                    .finally(() => this.$scope.resetAction());
            };
        }
    });
