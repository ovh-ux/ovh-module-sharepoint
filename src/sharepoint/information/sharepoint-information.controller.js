angular
  .module('Module.sharepoint.controllers')
  .controller('SharepointInformationsCtrl', class SharepointInformationsCtrl {
    constructor(
      $scope, $location, $stateParams,
      Alerter, MicrosoftSharepointLicenseService, Products,
    ) {
      this.$scope = $scope;
      this.$location = $location;
      this.$stateParams = $stateParams;
      this.alerter = Alerter;
      this.sharepointService = MicrosoftSharepointLicenseService;
      this.productsService = Products;
    }

    $onInit() {
      this.loaders = {
        init: false,
      };

      this.getProductsByType();
      this.getSharepoint();
      this.getExchangeOrganization();
    }

    getProductsByType() {
      return this.productsService.getProductsByType()
        .then((products) => {
          this.associatedExchange = _.find(products.exchanges, {
            name: this.$stateParams.exchangeId,
          });
          if (this.associatedExchange) {
            this.associatedExchangeLink = `#/configuration/${this.associatedExchange.type.toLowerCase()}/${this.associatedExchange.organization}/${this.associatedExchange.name}`;
          }
        });
    }

    getSharepoint() {
      this.loaders.init = true;
      return this.sharepointService.getSharepoint(this.$stateParams.exchangeId)
        .then((sharepoint) => {
          this.sharepoint = sharepoint;
          if (!this.sharepoint.url) {
            this.$location.path(`/configuration/sharepoint/${this.$stateParams.exchangeId}/${this.sharepoint.domain}/setUrl`);
          } else {
            this.calculateQuotas(sharepoint);
          }
        })
        .catch((err) => {
          _.set(err, 'type', err.type || 'ERROR');
          this.alerter.alertFromSWS(this.$scope.tr('sharepoint_dashboard_error'), err, this.$scope.alerts.main);
        })
        .finally(() => {
          this.loaders.init = false;
        });
    }

    getExchangeOrganization() {
      return this.sharepointService.retrievingExchangeOrganization(this.$stateParams.exchangeId)
        .then((organization) => {
          this.hideAssociatedExchange = !organization;
        });
    }

    calculateQuotas(sharepoint) {
      if (sharepoint.quota && sharepoint.currentUsage) {
        _.set(sharepoint, 'left', sharepoint.quota - sharepoint.currentUsage);
      }
      this.sharepoint = sharepoint;
    }

    setExchange() {
      this.productsService.setSelectedProduct(this.associatedExchange);
    }
  });
