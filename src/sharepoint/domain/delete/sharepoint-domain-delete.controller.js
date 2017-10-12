angular
    .module("Module.sharepoint.controllers")
    .controller("SharepointDeleteDomainController", class SharepointDeleteDomainController {

        constructor ($scope, $stateParams, Alerter, MicrosoftSharepointLicenseService) {
            this.$scope = $scope;
            this.$stateParams = $stateParams;
            this.Alerter = Alerter;
            this.SharepointService = MicrosoftSharepointLicenseService;
        }

        $onInit () {
            this.domain = this.$scope.currentActionData;
            this.$scope.deleteDomain = () => this.deleteDomain();
        }

        deleteDomain () {
            return this.SharepointService.deleteSharepointUpnSuffix(this.$stateParams.exchangeId, this.domain.suffix)
                .then(() => {
                    this.Alerter.success(this.$scope.tr("sharepoint_delete_domain_confirm_message_text", [this.domain.suffix]), this.$scope.alerts.main);

                    // TODO refresh domain's table
                })
                .catch((err) => {
                    this.Alerter.alertFromSWS(this.$scope.tr("sharepoint_delete_domain_error_message_text"), err, this.$scope.alerts.main);
                })
                .finally(() => {
                    this.$scope.resetAction();
                });
        }
    });
