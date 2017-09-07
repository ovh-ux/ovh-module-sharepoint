/* eslint-disable no-shadow, class-methods-use-this */
angular
    .module("Module.sharepoint.controllers")
    .controller("SharepointOrderCtrl", class SharepointOrderCtrl {

        constructor (constants, Exchange, MicrosoftSharepointLicenseService, Products, $q, $scope, $stateParams, User) {
            this.constants = constants;
            this.exchangeService = Exchange;
            this.sharepointService = MicrosoftSharepointLicenseService;
            this.productsService = Products;
            this.$q = $q;
            this.$scope = $scope;
            this.$stateParams = $stateParams;
            this.userService = User;
        }

        $onInit () {
            this.organizationId = this.$stateParams.organizationId;
            this.exchangeId = this.$stateParams.exchangeId;

            this.alerts = {
                dashboard: "sharepointDashboardAlert"
            };

            this.associatedExchange = null;
            this.isCommingFromAssociatedExchange = false; // Indicates that we are activating a SharePoint by coming from an exchange service.
            this.loaders = {
                init: true
            };
            this.worldPart = this.constants.target;
            this.associateExchange = true;
            this.standAloneQuantity = 1;
            this.accountsToActivate = [];

            this.getExchanges()
                .finally(() => {
                    this.loaders.init = false;
                    this.associatedExchange = this.exchanges[0];
                    this.getAccounts();
                });

            this.userService.getUser()
                .then((user) => { this.userSubsidiary = user.ovhSubsidiary; });
        }

        getExchanges () {
            return this.productsService.getProductsByType()
                .then((productsByType) => productsByType.exchanges)
                .then((exchanges) =>

                    // The Exchange service expects the exchange object to have a domain property.
                    // This property is returned by a call to /email/exchange/{organizationName}/service/{exchangeService}
                    // As far as I could tell, it is always the same thing as the name (exchangeService).
                    // So I deciced to hack a domain property onto the exchange object instead of calling the API.
                    _.map(exchanges, (exchange) => {
                        const newExchange = angular.copy(exchange);

                        newExchange.domain = newExchange.name;
                        return newExchange;
                    }))
                .then((exchanges) => this.filterSupportedExchanges(exchanges))
                .then((exchanges) => {
                    this.exchanges = exchanges;
                    if (exchanges.length === 0) {
                        this.associateExchange = false;
                    }
                    return exchanges;
                })
                .then((exchanges) => this.selectAssociatedExchangeFromRouteParams(exchanges));
        }

        filterSupportedExchanges (exchanges) {
            return _(exchanges)
                .filter((exchange) => this.isSupportedExchangeType(exchange))
                .filter((exchange) => this.isSupportedExchangeAdditionalCondition(exchange))
                .thru((exchanges) => this.filterExchangesThatAlreadyHaveSharepoint(exchanges))
                .value();
        }

        isSupportedExchangeType (exchange) {
            return exchange.type === "EXCHANGE_HOSTED";
        }

        isSupportedExchangeAdditionalCondition (exchange) {
            return this.exchangeService.getExchangeServer(exchange.organization, exchange.name)
                .then((server) => server.individual2010 === false);
        }

        filterExchangesThatAlreadyHaveSharepoint (exchanges) {
            return this.buildSharepointChecklist(exchanges)
                .then((checklist) => _.filter(exchanges, (exchange, index) => !checklist[index]));
        }

        selectAssociatedExchangeFromRouteParams (exchanges) {
            this.associatedExchange = _.find(exchanges, (exchange) => exchange.organization === this.organizationId && exchange.name === this.exchangeId);
            if (this.associatedExchange) {
                this.isCommingFromAssociatedExchange = true;
            }
        }

        buildSharepointChecklist (exchanges) {
            return this.$q.all(
                _.map(exchanges, (exchange) => this.hasSharepoint(exchange))
            );
        }

        hasSharepoint (exchange) {
            return this.exchangeService.getSharepointServiceForExchange(exchange)
                .then(() => true)
                .catch(() => false);
        }

        getAccounts () {
            this.loaders.accounts = true;
            this.accountsIds = null;

            return this.exchangeService.getAccountIds({
                organizationName: this.associatedExchange.organization,
                exchangeService: this.associatedExchange.domain
            }).then((accounts) => {
                this.accountsIds = accounts;
            }).finally(() => {
                if (_.isEmpty(this.accountsIds)) {
                    this.loaders.accounts = false;
                }
            });
        }

        onTranformItem (account) {
            return this.exchangeService.getAccount({
                organizationName: this.associatedExchange.organization,
                exchangeService: this.associatedExchange.domain,
                primaryEmailAddress: account
            })
                .then((account) => {
                    account.activateSharepoint = _.some(this.accountsToActivate, account.primaryEmailAddress);
                    return account;
                });
        }

        onTranformItemDone () {
            this.loaders.accounts = false;
        }

        checkSharepointActivation (account) {
            if (account.activateSharepoint) {
                this.accountsToActivate.push(account.primaryEmailAddress);
            } else {
                this.accountsToActivate.splice(this.accountsToActivate.indexOf(account.primaryEmailAddress), 1);
            }
        }

        getExchangeAccountsUrl () {
            return `#/configuration/exchange_hosted/${this.associatedExchange.organization}/${this.associatedExchange.name}?tab=ACCOUNTS`;
        }

        hasInexistantAccount (activateSharepointForm) {
            return activateSharepointForm.primaryEmailAddressField.$invalid && !_.isEmpty(activateSharepointForm.primaryEmailAddressField.$viewValue);
        }

        changeStandAloneQuantity () {
            if (parseInt(this.standAloneQuantity, 10) >= 1 && parseInt(this.standAloneQuantity, 10) <= 30) {
                this.standAloneQuantity = parseInt(this.standAloneQuantity, 10);
            }
        }

        decrement () {
            if (this.standAloneQuantity > 1) {
                this.standAloneQuantity--;
            }
        }

        increment () {
            if (this.standAloneQuantity < 30) {
                this.standAloneQuantity++;
            }
        }

        getSharepointOrderUrl () {
            if (this.associateExchange) {
                if (_.has(this.associatedExchange, "name") && this.accountsToActivate.length >= 1) {
                    return this.sharepointService.getSharepointOrderUrl(this.associatedExchange.name, this.accountsToActivate);
                }
                return "";

            } else if (!_.isNull(this.standAloneQuantity) && parseInt(this.standAloneQuantity, 10) >= 1 && parseInt(this.standAloneQuantity, 10) <= 30) {
                return this.sharepointService.getSharepointStandaloneOrderUrl(parseInt(this.standAloneQuantity, 10));
            }
            return "";

        }
    });
/* eslint-enable no-shadow, class-methods-use-this */
