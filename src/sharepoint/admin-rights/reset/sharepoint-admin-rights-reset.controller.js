angular
    .module("Module.sharepoint.controllers")
    .controller("SharepointResetAdminRightsCtrl", class SharepointResetAdminRightsCtrl {

        constructor (Alerter, MicrosoftSharepointLicenseService, $stateParams, $scope) {
            this.alerter = Alerter;
            this.sharepointService = MicrosoftSharepointLicenseService;
            this.$stateParams = $stateParams;
            this.$scope = $scope;
        }

        $onInit () {
            this.sharepointDomain = this.$stateParams.productId;
            this.exchangeId = this.$stateParams.exchangeId;

            this.$scope.submit = () => {
                this.$scope.resetAction();

                return this.sharepointService.restoreAdminRights({
                    serviceName: this.exchangeId
                })
                    .then(() => this.alerter.success(this.$scope.tr("sharepoint_accounts_action_reset_admin_success"), this.$scope.alerts.dashboard))
                    .catch((err) => this.alerter.alertFromSWS(this.$scope.tr("sharepoint_accounts_action_reset_admin_error"), err, this.$scope.alerts.dashboard))
                    .finally(() => this.$scope.resetAction());
            };
        }
    });
