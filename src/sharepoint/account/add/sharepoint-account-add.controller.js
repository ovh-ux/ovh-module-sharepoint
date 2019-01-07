angular
  .module('Module.sharepoint.controllers')
  .controller('SharepointAccountAddCtrl', class SharepointAccountAddCtrl {
    constructor(
      $scope,
      $stateParams,
      $translate,
      MicrosoftSharepointOrderService,
    ) {
      this.$scope = $scope;
      this.$stateParams = $stateParams;
      this.$translate = $translate;
      this.sharepointOrder = MicrosoftSharepointOrderService;
    }

    $onInit() {
      this.loading = true;
      this.prices = null;
      this.sharepointOrder.creatingCart()
        .then(cartId => this.sharepointOrder.fetchingPrices(cartId))
        .then((prices) => {
          this.price = prices.get('P1M');
          this.loading = false;
        });
    }

    getPriceText(quantity) {
      return `${this.price.value * quantity} ${this.price.currencyCode === 'EUR' ? '&#0128;' : this.price.currencyCode}`;
    }

    submit() {
      this.alerter.success(this.$translate.instant('sharepoint_account_action_sharepoint_add_success_message'), this.$scope.alerts.main);
      this.$scope.resetAction();
      window.open(this.sharepointService.getSharepointStandaloneNewAccountOrderUrl(
        this.$stateParams.productId,
        this.optionsList[0].prices[0].quantity,
      ));
    }
  });
