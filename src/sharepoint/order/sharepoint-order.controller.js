angular
  .module('Module.sharepoint.controllers')
  .controller('SharepointOrderCtrl', class SharepointOrderCtrl {
    constructor(
      $scope, $q, $stateParams,
      constants, Exchange, MicrosoftSharepointLicenseService, User,
    ) {
      this.$scope = $scope;
      this.$q = $q;
      this.$stateParams = $stateParams;
      this.constants = constants;
      this.exchangeService = Exchange;
      this.sharepointService = MicrosoftSharepointLicenseService;
      this.userService = User;
    }

    $onInit() {
      this.organizationId = this.$stateParams.organizationId;
      this.exchangeId = this.$stateParams.exchangeId;

      this.alerts = {
        main: 'sharepoint.alerts.main',
      };

      this.associatedExchange = null;
      // Indicates that we are activating a SharePoint by coming from an exchange service.
      this.isCommingFromAssociatedExchange = false;
      this.loaders = {
        init: true,
      };
      this.worldPart = this.constants.target;
      this.associateExchange = true;
      this.standAloneQuantity = 1;
      this.accountsToActivate = [];

      this.getExchanges()
        .finally(() => {
          this.loaders.init = false;
          this.associatedExchange = _.first(this.exchanges);
          this.getAccounts();
        });

      this.userService.getUser()
        .then((user) => { this.userSubsidiary = user.ovhSubsidiary; });
    }

    getExchanges() {
      return this.sharepointService.getExchangeServices()
        .then(exchanges => _.map(exchanges, (exchange) => {
          const newExchange = angular.copy(exchange);

          newExchange.domain = newExchange.name;
          return newExchange;
        }))
        .then(exchanges => this.filterSupportedExchanges(exchanges))
        .then((exchanges) => {
          this.exchanges = exchanges;
          if (exchanges.length === 0) {
            this.associateExchange = false;
          }
          return exchanges;
        })
        .then(exchanges => this.selectAssociatedExchangeFromRouteParams(exchanges));
    }

    filterSupportedExchanges(exchanges) {
      return _(exchanges)
        .filter(exchange => this.constructor.isSupportedExchangeType(exchange))
        .filter(exchange => this.isSupportedExchangeAdditionalCondition(exchange))
        .thru(exchanges => this.filterExchangesThatAlreadyHaveSharepoint(exchanges)) // eslint-disable-line
        .value();
    }

    static isSupportedExchangeType(exchange) {
      return exchange.type === 'EXCHANGE_HOSTED';
    }

    isSupportedExchangeAdditionalCondition(exchange) {
      return this.exchangeService.getExchangeServer(exchange.organization, exchange.name)
        .then(server => server.individual2010 === false);
    }

    filterExchangesThatAlreadyHaveSharepoint(exchanges) {
      return this.buildSharepointChecklist(exchanges)
        .then(checklist => _.filter(exchanges, (exchange, index) => !checklist[index]));
    }

    selectAssociatedExchangeFromRouteParams(exchanges) {
      this.associatedExchange = _.find(
        exchanges,
        exchange => exchange.organization === this.organizationId
          && exchange.name === this.exchangeId,
      );
      if (this.associatedExchange) {
        this.isCommingFromAssociatedExchange = true;
      }
    }

    buildSharepointChecklist(exchanges) {
      return this.$q.all(
        _.map(exchanges, exchange => this.hasSharepoint(exchange)),
      );
    }

    hasSharepoint(exchange) {
      return this.exchangeService.getSharepointServiceForExchange(exchange)
        .then(() => true)
        .catch(() => false);
    }

    getAccounts() {
      this.loaders.accounts = true;
      this.accountsIds = null;

      return this.exchangeService.getAccountIds({
        organizationName: this.associatedExchange.organization,
        exchangeService: this.associatedExchange.domain,
      }).then((accounts) => {
        this.accountsIds = accounts;
      }).finally(() => {
        if (_.isEmpty(this.accountsIds)) {
          this.loaders.accounts = false;
        }
      });
    }

    onTranformItem(account) {
      return this.exchangeService
        .getAccount({
          organizationName: this.associatedExchange.organization,
          exchangeService: this.associatedExchange.domain,
          primaryEmailAddress: account,
        })
        .then((account) => { // eslint-disable-line
          _.set(account, 'activateSharepoint', _.some(this.accountsToActivate, account.primaryEmailAddress));
          return account;
        });
    }

    onTranformItemDone() {
      this.loaders.accounts = false;
    }

    checkSharepointActivation(account) {
      if (account.activateSharepoint) {
        this.accountsToActivate.push(account.primaryEmailAddress);
      } else {
        this.accountsToActivate
          .splice(this.accountsToActivate.indexOf(account.primaryEmailAddress), 1);
      }
    }

    getExchangeAccountsUrl() {
      return `#/configuration/exchange_hosted/${this.associatedExchange.organization}/${this.associatedExchange.name}?tab=ACCOUNTS`;
    }

    static hasInexistantAccount(activateSharepointForm) {
      return activateSharepointForm.primaryEmailAddressField.$invalid
        && !_.isEmpty(activateSharepointForm.primaryEmailAddressField.$viewValue);
    }

    changeStandAloneQuantity() {
      if (parseInt(this.standAloneQuantity, 10) >= 1
        && parseInt(this.standAloneQuantity, 10) <= 30) {
        this.standAloneQuantity = parseInt(this.standAloneQuantity, 10);
      }
    }

    decrement() {
      if (this.standAloneQuantity > 1) {
        this.standAloneQuantity -= 1;
      }
    }

    increment() {
      if (this.standAloneQuantity < 30) {
        this.standAloneQuantity += 1;
      }
    }

    getSharepointOrderUrl() {
      if (this.associateExchange) {
        if (_.has(this.associatedExchange, 'name') && this.accountsToActivate.length >= 1) {
          return this.sharepointService.getSharepointOrderUrl(
            this.associatedExchange.name,
            this.accountsToActivate,
          );
        }
        return '';
      }
      if (!_.isNull(this.standAloneQuantity)
        && parseInt(this.standAloneQuantity, 10) >= 1
        && parseInt(this.standAloneQuantity, 10) <= 30) {
        return this.sharepointService
          .getSharepointStandaloneOrderUrl(parseInt(this.standAloneQuantity, 10));
      }

      return '';
    }
  });
