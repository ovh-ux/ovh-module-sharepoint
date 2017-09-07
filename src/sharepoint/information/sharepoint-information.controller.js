angular
    .module("Module.sharepoint.controllers")
    .controller("SharepointInformationsCtrl", class SharepointInformationsCtrl {

        constructor (Alerter, $location, MicrosoftSharepointLicenseService, Products, $stateParams, $scope) {
            this.alerter = Alerter;
            this.$location = $location;
            this.sharepointService = MicrosoftSharepointLicenseService;
            this.productsService = Products;
            this.$stateParams = $stateParams;
            this.$scope = $scope;
        }

        $onInit () {
            this.accountIds = [];
            this.loaders = {
                init: false
            };

            this.getProductsByType();
            this.getSharepoint();
            this.getExchangeOrganization();
        }

        getProductsByType () {
            return this.productsService.getProductsByType()
                .then((products) => {
                    this.associatedExchange = _.find(products.exchanges, { name: this.$stateParams.exchangeId });
                    if (this.associatedExchange) {
                        this.associatedExchangeLink = ["#/configuration", this.associatedExchange.type.toLowerCase(),
                            this.associatedExchange.organization, this.associatedExchange.name].join("/");
                    }
                });
        }

        getSharepoint () {
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
                    this.alerter.alertFromSWS(this.$scope.tr("sharepoint_dashboard_error"), err);
                })
                .finally(() => {
                    this.loaders.init = false;
                });
        }

        getExchangeOrganization () {
            return this.sharepointService.retrievingExchangeOrganization(this.$stateParams.exchangeId)
                .then((organization) => {
                    this.hideAssociatedExchange = !organization;
                });
        }

        calculateQuotas (sharepoint) {
            if (sharepoint.quota && sharepoint.currentUsage) {
                sharepoint.left = sharepoint.quota - sharepoint.currentUsage;
            }
            this.sharepoint = sharepoint;
        }

        onTranformItem (userPrincipalName) {
            this.loaders.init = true;
            return this.sharepointService.getAccount({
                organizationName: this.$stateParams.organization,
                sharepointService: this.$stateParams.productId,
                userPrincipalName
            })
                .then((account) => {
                    this.sharepointService.getAccountTasks({
                        organizationName: this.$stateParams.organization,
                        sharepointService: this.$stateParams.productId,
                        userPrincipalName
                    })
                        .then((tasks) => {
                            account.tasksCount = tasks.length;
                            return account;
                        });
                });
        }

        onTranformItemDone () {
            this.loaders.init = false;
        }

        setExchange () {
            this.productsService.setSelectedProduct(this.associatedExchange);
        }
    });
