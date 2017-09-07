angular
    .module("Module.sharepoint.controllers")
    .controller("SharepointAddDomainController", class SharepointAddDomainController {

        constructor (Alerter, MicrosoftSharepointLicenseService, Products, $stateParams, $scope, Validator) {
            this.alerter = Alerter;
            this.sharepointService = MicrosoftSharepointLicenseService;
            this.productsService = Products;
            this.$stateParams = $stateParams;
            this.$scope = $scope;
            this.validatorService = Validator;
        }

        $onInit () {
            this.punycode = window.punycode;

            this.loading = true;
            this.domainType = "ovhDomain";
            this.model = {
                name: "",
                displayName: null
            };

            this.$scope.loadDomainData = () => {
                this.loading = true;

                return this.productsService.getProductsByType()
                    .then((productsByType) => this.prepareData(productsByType.domains))
                    .catch((err) => this.alerter.alertFromSWS(this.$scope.tr("sharepoint_add_domain_error_message_text"), err, this.$scope.alerts.dashboard));
            };

            this.$scope.addDomain = () => this.sharepointService.addSharepointUpnSuffixe(this.$stateParams.exchangeId, this.model.name)
                .then(() => this.alerter.success(this.$scope.tr("sharepoint_add_domain_confirm_message_text", this.model.displayName), this.$scope.alerts.dashboard))
                .catch((err) => this.alerter.alertFromSWS(this.$scope.tr("sharepoint_add_domain_error_message_text"), err, this.$scope.alerts.dashboard))
                .finally(() => this.$scope.resetAction());
        }

        prepareData (data) {
            const domains = _.filter(data, (item) => item.type === "DOMAIN");

            return this.sharepointService.getUsedUpnSuffixes()
                .then((upnSuffixes) => {
                    _.remove(domains, (domain) => upnSuffixes.indexOf(domain.name) >= 0);
                })
                .finally(() => {
                    this.loading = false;
                    this.availableDomains = domains;
                    this.availableDomainsBuffer = _.clone(this.availableDomains);
                });
        }

        resetSearchValue () {
            this.search.value = null;
            this.availableDomains = _.clone(this.availableDomainsBuffer);
        }

        resetName () {
            if (!_.isUndefined(this.search) && _.has(this.search, "value")) {
                this.availableDomains = _.filter(this.availableDomainsBuffer, (n) => n.displayName.search(this.search.value) > -1);
            }

            this.model.displayName = null;
            this.model.name = "";
        }

        changeName () {
            this.model.name = this.punycode.toASCII(this.model.displayName);

            // URL validation based on http://www.regexr.com/39nr7
            this.domainValid = this.validatorService.isValidURL(this.model.name);
        }
    });
