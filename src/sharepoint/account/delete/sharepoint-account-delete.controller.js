angular
    .module("Module.sharepoint.controllers")
    .controller("SharepointDeleteAccountCtrl", class SharepointDeleteAccountCtrl {

        constructor ($scope, $stateParams, Alerter, MicrosoftSharepointLicenseService) {
            this.$scope = $scope;
            this.$stateParams = $stateParams;
            this.Alerter = Alerter;
            this.SharepointService = MicrosoftSharepointLicenseService;
        }

        $onInit () {
            this.account = this.$scope.currentActionData;
            this.$scope.submit = () => this.submit();
        }

        submit () {
            this.$scope.resetAction();
            return this.SharepointService.deleteSharepointAccount(this.$stateParams.exchangeId, this.account.userPrincipalName)
                .then(() => {
                    this.Alerter.success(this.$scope.tr("sharepoint_account_action_sharepoint_remove_success_message", this.account.userPrincipalName), this.$scope.alerts.main);
                })
                .catch((err) => {
                    this.Alerter.alertFromSWS(this.$scope.tr("sharepoint_account_action_sharepoint_remove_error_message"), err, this.$scope.alerts.main);
                })
                .finally(() => {
                    this.$scope.resetAction();
                });
        }
    });
