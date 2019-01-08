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
      this.price = null;
      this.quantity = 1;
      this.sharepointOrder.creatingCart()
        .then(cartId => this.sharepointOrder.fetchingPrices(
          cartId,
          'activedirectory-account-provider',
          'sharepoint-account-provider-2016',
        ))
        .then((prices) => {
          this.price = _.get(prices, 'P1M');
          this.loading = false;
        });
    }

    getTotalPrice() {
      return _.round(_.get(this.price, 'value', 0) * this.quantity, 2);
    }

    getCurrency() {
      return _.get(this.price, 'currencyCode') === 'EUR' ? '&#0128;' : _.get(this.price, 'currencyCode');
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
