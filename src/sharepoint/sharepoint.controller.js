angular
    .module("Module.sharepoint.controllers")
    .controller("SharepointCtrl", class SharepointCtrl {

        constructor (Alerter, constants, MicrosoftSharepointLicenseService, Products, $stateParams, $scope, $timeout) {
            this.alerter = Alerter;
            this.constants = constants;
            this.sharepointService = MicrosoftSharepointLicenseService;
            this.productsService = Products;
            this.$stateParams = $stateParams;
            this.$scope = $scope;
            this.$timeout = $timeout;
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
                dashboard: "sharepointDashboardAlert"
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
                .then((sharepoint) => { this.sharepoint = sharepoint; })
                .catch((err) => this.alerter.alertFromSWS(this.$scope.tr("sharepoint_dashboard_error"), err))
                .finally(() => { this.loaders.init = false; });
        }

        getProducts () {
            return this.productsService.getProducts()
                .then((products) => {
                    let exchange = _.find(products, { name: this.exchangeId });

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
