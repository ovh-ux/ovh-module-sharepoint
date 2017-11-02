angular
    .module("Module.sharepoint.controllers")
    .controller("SharepointDeleteDomainController", class SharepointDeleteDomainController {

        constructor ($scope, $stateParams, Alerter, MicrosoftSharepointLicenseService) {
            this.$scope = $scope;
            this.$stateParams = $stateParams;
            this.alerter = Alerter;
            this.sharepointService = MicrosoftSharepointLicenseService;
        }

        $onInit () {
            this.domain = this.$scope.currentActionData;
            this.$scope.deleteDomain = () => this.deleteDomain();
        }

        deleteDomain () {
            return this.sharepointService.deleteSharepointUpnSuffix(this.$stateParams.exchangeId, this.domain.suffix)
                .then(() => {
                    this.alerter.success(this.$scope.tr("sharepoint_delete_domain_confirm_message_text", [this.domain.suffix]), this.$scope.alerts.main);

                    // TODO refresh domain's table
                })
                .catch((err) => {
                    this.alerter.alertFromSWS(this.$scope.tr("sharepoint_delete_domain_error_message_text"), err, this.$scope.alerts.main);
                })
                .finally(() => {
                    this.$scope.resetAction();
                });
        }
    });
