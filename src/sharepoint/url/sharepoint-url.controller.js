angular
  .module('Module.sharepoint.controllers')
  .controller('SharepointUrlCtrl', class SharepointUrlCtrl {
    constructor(
      $location, $scope, $stateParams, $timeout, $translate,
      Alerter, constants, MicrosoftSharepointLicenseService,
    ) {
      this.$location = $location;
      this.$scope = $scope;
      this.$stateParams = $stateParams;
      this.$timeout = $timeout;
      this.$translate = $translate;
      this.alerter = Alerter;
      this.constants = constants;
      this.sharepointService = MicrosoftSharepointLicenseService;
    }

    $onInit() {
      this.alerts = {
        main: 'sharepoint.alerts.main',
      };

      this.sharepointDomain = this.$stateParams.productId;
      this.exchangeId = this.$stateParams.exchangeId;
      this.loaders = {
        init: false,
      };

      this.worldPart = this.constants.target;

      this.sharepointUrl = null;

      this.retrievingSharepointSuffix();
      this.assignGuideUrl();
    }

    retrievingSharepointSuffix() {
      return this.sharepointService
        .retrievingSharepointSuffix(this.exchangeId)
        .then((suffix) => {
          this.sharepointUrlSuffix = suffix;
        });
    }

    assignGuideUrl() {
      return this.sharepointService.assignGuideUrl(this, 'sharepointGuideUrl');
    }

    activatingSharepoint() {
      return this.sharepointService
        .setSharepointUrl(this.exchangeId, `${this.sharepointUrl}${this.sharepointUrlSuffix}`)
        .then(() => {
          this.alerter.success(this.$translate.instant('sharepoint_set_url_success_message_text', { t0: this.exchangeId }), this.alerts.main);

          this.$timeout(() => {
            this.$location.path(`/configuration/sharepoint/${this.exchangeId}/${this.sharepointDomain}`);
          }, 3000);
        })
        .catch((err) => {
          _.set(err, 'type', err.type || 'ERROR');
          this.alerter.alertFromSWS(this.$translate.instant('sharepoint_set_url_failure_message_text'), err, this.alerts.main);
        });
    }
  });
