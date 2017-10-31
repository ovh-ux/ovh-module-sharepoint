angular
    .module("Module.sharepoint.controllers")
    .controller("SharepointUrlCtrl", class SharepointUrlCtrl {

        constructor ($location, $scope, $stateParams, $timeout, Alerter, constants, MicrosoftSharepointLicenseService) {
            this.alerter = Alerter;
            this.constants = constants;
            this.$location = $location;
            this.sharepointService = MicrosoftSharepointLicenseService;
            this.$scope = $scope;
            this.$stateParams = $stateParams;
            this.$timeout = $timeout;
        }

        $onInit () {
            this.$scope.alerts = {
                page: "sharepoint.alerts.page",
                main: "sharepoint.alerts.main"
            };

            this.sharepointDomain = this.$stateParams.productId;
            this.exchangeId = this.$stateParams.exchangeId;
            this.loaders = {
                init: false
            };

            this.worldPart = this.constants.target;

            this.sharepointUrl = null;

            this.retrievingSharepointSuffix();
            this.assignGuideUrl();
        }

        retrievingSharepointSuffix () {
            return this.sharepointService
                .retrievingSharepointSuffix(this.exchangeId)
                .then((suffix) => {
                    this.sharepointUrlSuffix = suffix;
                });
        }

        assignGuideUrl () {
            return this.sharepointService.assignGuideUrl(this, "sharepointGuideUrl");
        }

        activatingSharepoint () {
            return this.sharepointService
                .setSharepointUrl(this.exchangeId, `${this.sharepointUrl}${this.sharepointUrlSuffix}`)
                .then(() => {
                    this.alerter.success(this.$scope.tr("sharepoint_set_url_success_message_text", this.exchangeId), this.$scope.alerts.main);

                    this.$timeout(() => {
                        this.$location.path(`/configuration/sharepoint/${this.exchangeId}/${this.sharepointDomain}`);
                    }, 3000);
                })
                .catch((err) => {
                    _.set(err, "type", err.type || "ERROR");
                    this.alerter.alertFromSWS(this.$scope.tr("sharepoint_set_url_failure_message_text"), err, this.$scope.alerts.main);
                });
        }
    });
