angular
  .module('Module.sharepoint.controllers')
  .controller('SharepointOrderCtrl', class SharepointOrderCtrl {
    constructor(
      $q,
      $stateParams,
      Exchange,
      MicrosoftSharepointLicenseService,
      ouiDatagridService,
      User,
    ) {
      this.$q = $q;
      this.$stateParams = $stateParams;
      this.Exchange = Exchange;
      this.Sharepoint = MicrosoftSharepointLicenseService;
      this.ouiDatagridService = ouiDatagridService;
      this.User = User;
    }

    $onInit() {
      this.organizationId = this.$stateParams.organizationId;
      this.exchangeId = this.$stateParams.exchangeId;

      this.alerts = {
        main: 'sharepoint.alerts.main',
      };

      this.canLinkToExchange = true;
      this.associateExchange = false;
      this.associatedExchange = null;

      // Indicates that we are activating a SharePoint by coming from an exchange service.
      this.isComingFromAssociatedExchange = false;

      this.loaders = {
        init: true,
        accounts: true,
      };

      this.standAloneQuantity = 1;
      this.accountsToActivate = [];

      this.getExchanges()
        .finally(() => {
          this.loaders.init = false;
          this.associatedExchange = this.associatedExchange || _.first(this.exchanges);
          this.getAccounts();
        });

      this.User.getUser()
        .then((user) => { this.userSubsidiary = user.ovhSubsidiary; });
    }

    getExchanges() {
      return this.Sharepoint.getExchangeServices()
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
      return this.Exchange.getExchangeServer(exchange.organization, exchange.name)
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
        this.isComingFromAssociatedExchange = true;
      }
    }

    buildSharepointChecklist(exchanges) {
      return this.$q.all(
        _.map(exchanges, exchange => this.hasSharepoint(exchange)),
      );
    }

    hasSharepoint(exchange) {
      return this.Exchange.getSharepointServiceForExchange(exchange)
        .then(() => true)
        .catch(() => false);
    }

    getAccounts() {
      return this.Exchange.getAccountIds({
        organizationName: this.associatedExchange.organization,
        exchangeService: this.associatedExchange.domain,
      }).then(accountEmails => ({
        data: _.map(accountEmails, email => ({ email })),
        meta: {
          totalCount: accountEmails.length,
        },
      }));
    }

    getAccount(row) {
      return this.Exchange
        .getAccount({
          organizationName: this.associatedExchange.organization,
          exchangeService: this.associatedExchange.domain,
          primaryEmailAddress: row.email,
        })
        .then((account) => {
          _.set(account, 'activateSharepoint', _.some(this.accountsToActivate, account.primaryEmailAddress));
          return account;
        });
    }

    refreshAccounts(exchange) {
      this.associatedExchange = exchange;
      this.ouiDatagridService.refresh('exchangeAccountsDatagrid');
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

    getSharepointOrderUrl() {
      if (this.associateExchange) {
        if (_.has(this.associatedExchange, 'name') && this.accountsToActivate.length >= 1) {
          return this.Sharepoint.getSharepointOrderUrl(
            this.associatedExchange.name,
            this.accountsToActivate,
          );
        }
        return '';
      }
      if (!_.isNull(this.standAloneQuantity)
        && parseInt(this.standAloneQuantity, 10) >= 1
        && parseInt(this.standAloneQuantity, 10) <= 30) {
        return this.Sharepoint
          .getSharepointStandaloneOrderUrl(parseInt(this.standAloneQuantity, 10));
      }

      return '';
    }

    goToSharepointOrder() {

    }
  });
