angular
    .module("Module.sharepoint.controllers")
    .controller("SharepointDeleteDomainController", class SharepointDeleteDomainController {

        constructor (Alerter, MicrosoftSharepointLicenseService, $stateParams, $scope) {
            this.alerter = Alerter;
            this.sharepointService = MicrosoftSharepointLicenseService;
            this.$stateParams = $stateParams;
            this.$scope = $scope;
        }

        $onInit () {
            this.domain = this.$scope.currentActionData;

            this.$scope.deleteDomain = () => this.sharepointService.deleteSharepointUpnSuffix(this.$stateParams.exchangeId, this.domain.suffix)
                .then(() => {
                    this.alerter.success(this.$scope.tr("sharepoint_delete_domain_confirm_message_text", [this.domain.suffix]), this.$scope.alerts.dashboard);

                    // TODO refresh domain's table
                })
                .catch((err) => this.alerter.alertFromSWS(this.$scope.tr("sharepoint_delete_domain_error_message_text"), err, this.$scope.alerts.dashboard))
                .finally(() => this.$scope.resetAction());
        }

    });
