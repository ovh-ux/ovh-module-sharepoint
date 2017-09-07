angular
    .module("Module.sharepoint.controllers")
    .controller("SharepointAddAccountCtrl", class SharepointAddAccountCtrl {

        constructor (Alerter, MicrosoftSharepointLicenseService, $stateParams, $scope) {
            this.alerter = Alerter;
            this.sharepointService = MicrosoftSharepointLicenseService;
            this.$stateParams = $stateParams;
            this.$scope = $scope;
        }

        $onInit () {
            this.loading = false;
            this.sharepointDomain = this.$stateParams.productId;

            this.getSherpointServiceOptions();

            this.$scope.submit = () => {
                this.alerter.success(this.$scope.tr("sharepoint_account_action_sharepoint_add_success_message"), this.$scope.alerts.dashboard);
                this.$scope.resetAction();
                window.open(
                    this.sharepointService.getSharepointStandaloneNewAccountOrderUrl(
                        this.sharepointDomain, this.optionsList[0].prices[0].quantity
                    )
                );
            };
        }

        getSherpointServiceOptions () {
            this.loading = true;

            return this.sharepointService.getSherpointServiceOptions(this.sharepointDomain)
                .then((options) => { this.optionsList = options; })
                .catch((err) => {
                    this.$scope.resetAction();
                    this.alerter.alertFromSWS(this.$scope.tr("sharepoint_accounts_action_sharepoint_add_error_message"), err, this.$scope.alerts.dashboard);
                })
                .finally(() => { this.loading = false; });
        }

        /* eslint-disable class-methods-use-this */
        getPrice (option) {
            return _.round(_.get(option, "prices[0].price.value", 0) * _.get(option, "prices[0].quantity", 0), 2);
        }

        getCurrency (option) {
            return _.get(option, "prices[0].price.currencyCode") === "EUR" ? "&#0128;" : _.get(option, "prices[0].price.currencyCode");
        }
        /* eslint-enable class-methods-use-this */
    });
