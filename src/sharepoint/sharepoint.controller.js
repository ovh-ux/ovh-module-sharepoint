angular
    .module("Module.sharepoint.controllers")
    .controller("SharepointCtrl", class SharepointCtrl {

        constructor ($scope, $stateParams, $timeout, Alerter, constants, MicrosoftSharepointLicenseService, Products) {
            this.$scope = $scope;
            this.$stateParams = $stateParams;
            this.$timeout = $timeout;
            this.alerter = Alerter;
            this.constants = constants;
            this.sharepointService = MicrosoftSharepointLicenseService;
            this.productsService = Products;
        }

        $onInit () {
            this.sharepointDomain = this.$stateParams.productId;
            this.exchangeId = this.$stateParams.exchangeId;
            this.worldPart = this.constants.target;
            this.sharepointService.assignGuideUrl(this, "sharepointGuideUrl");
            this.loaders = {
                init: true
            };
            this.stepPath = "";

            this.$scope.alerts = {
                page: "sharepoint.alerts.page",
                tabs: "sharepoint.alerts.tabs",
                main: "sharepoint.alerts.main"
            };

            this.getSharepoint();
            this.getProducts();
            this.getExchangeOrganization();

            this.$scope.currentAction = null;
            this.$scope.currentActionData = null;

            this.$scope.setAction = (action, data) => {
                this.$scope.currentAction = action;
                this.$scope.currentActionData = data;
                if (action) {
                    this.stepPath = `sharepoint/${this.$scope.currentAction}.html`;
                    $("#currentAction").modal({
                        keyboard: true,
                        backdrop: "static"
                    });
                } else {
                    $("#currentAction").modal("hide");
                    this.$scope.currentActionData = null;
                    this.$timeout(() => {
                        this.stepPath = "";
                    }, 300);
                }
            };

            this.$scope.resetAction = () => {
                this.$scope.setAction(false);
            };

            this.$scope.$on("$locationChangeStart", () => {
                this.$scope.resetAction();
            });
        }

        getSharepoint () {
            return this.sharepointService.getSharepoint(this.$stateParams.exchangeId)
                .then((sharepoint) => {
                    this.sharepoint = sharepoint;
                })
                .catch((err) => {
                    _.set(err, "type", err.type || "ERROR");
                    this.alerter.alertFromSWS(this.$scope.tr("sharepoint_dashboard_error"), err, this.$scope.alerts.page);
                })
                .finally(() => {
                    this.loaders.init = false;
                });
        }

        getProducts () {
            return this.productsService.getProducts()
                .then((products) => {
                    const exchange = _.find(products, { name: this.exchangeId });

                    if (exchange) {
                        this.exchangeOrganization = exchange.organization;
                    }
                });
        }

        getExchangeOrganization () {
            return this.sharepointService.retrievingExchangeOrganization(this.exchangeId)
                .then((organization) => {
                    if (!organization) {
                        this.isStandAlone = true;
                    }
                });
        }
    });
