angular
    .module("Module.sharepoint.controllers")
    .controller("SharepointActivateAccountCtrl", class SharepointActivateAccountCtrl {

        constructor (Alerter, MicrosoftSharepointLicenseService, $stateParams, $scope, User) {
            this.alerter = Alerter;
            this.sharepointService = MicrosoftSharepointLicenseService;
            this.$stateParams = $stateParams;
            this.$scope = $scope;
            this.userService = User;
        }

        $onInit () {
            this.account = this.$scope.currentActionData;

            this.$scope.submit = () => {
                this.$scope.resetAction();
                return this.sharepointService.activateSharepointOnAccount(this.$stateParams.exchangeId, this.account.userPrincipalName)
                    .then(() => {
                        this.alerter.success(
                            this.$scope.tr("sharepoint_action_activate_account_success_message", this.account.userPrincipalName),
                            this.$scope.alerts.dashboard
                        );
                    })
                    .catch((err) => this.alerter.alertFromSWS(this.$scope.tr("sharepoint_action_activate_account_error_message"), err, this.$scope.alerts.dashboard))
                    .finally(() => this.$scope.resetAction());
            };
        }

    });
