angular
  .module('Module.sharepoint.controllers')
  .controller('SharepointDeleteAccountCtrl', class SharepointDeleteAccountCtrl {
    constructor($scope, $stateParams, Alerter, MicrosoftSharepointLicenseService) {
      this.$scope = $scope;
      this.$stateParams = $stateParams;
      this.alerter = Alerter;
      this.sharepointService = MicrosoftSharepointLicenseService;
    }

    $onInit() {
      this.account = this.$scope.currentActionData;
      this.$scope.submit = () => this.submit();
    }

    submit() {
      this.$scope.resetAction();
      return this.sharepointService
        .deleteSharepointAccount(this.$stateParams.exchangeId, this.account.userPrincipalName)
        .then(() => this.alerter.success(this.$scope.tr('sharepoint_account_action_sharepoint_remove_success_message', this.account.userPrincipalName), this.$scope.alerts.main))
        .catch(err => this.alerter.alertFromSWS(this.$scope.tr('sharepoint_account_action_sharepoint_remove_error_message'), err, this.$scope.alerts.main))
        .finally(() => {
          this.$scope.resetAction();
        });
    }
  });
