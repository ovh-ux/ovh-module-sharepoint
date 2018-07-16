angular
  .module('Module.sharepoint.controllers')
  .controller('SharepointResetAdminRightsCtrl', class SharepointResetAdminRightsCtrl {
    constructor($scope, $stateParams, Alerter, MicrosoftSharepointLicenseService) {
      this.$scope = $scope;
      this.$stateParams = $stateParams;
      this.alerter = Alerter;
      this.sharepointService = MicrosoftSharepointLicenseService;
    }

    $onInit() {
      this.sharepointDomain = this.$stateParams.productId;
      this.exchangeId = this.$stateParams.exchangeId;

      this.$scope.submit = () => this.submit();
    }

    submit() {
      this.$scope.resetAction();
      return this.sharepointService.restoreAdminRights(this.exchangeId)
        .then(() => {
          this.alerter.success(this.$scope.tr('sharepoint_accounts_action_reset_admin_success'), this.$scope.alerts.main);
        })
        .catch((err) => {
          this.alerter.alertFromSWS(this.$scope.tr('sharepoint_accounts_action_reset_admin_error'), err, this.$scope.alerts.main);
        })
        .finally(() => {
          this.$scope.resetAction();
        });
    }
  });
