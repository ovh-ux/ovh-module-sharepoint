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
                dashboard: "sharepointDashboardAlert"
            };
            this.sharepointDomain = this.$stateParams.productId;
            this.exchangeId = this.$stateParams.exchangeId;
            this.loaders = {
                init: false
            };
            this.worldPart = this.constants.target;

            this.sharepointUrl = null;

            this.getSharepointSuffix();
            this.assignGuideUrl();
        }

        getSharepointSuffix () {
            return this.sharepointService.getSharepointSuffix(this.exchangeId)
                .then((suffix) => { this.sharepointUrlSuffix = suffix; });
        }

        assignGuideUrl () {
            return this.sharepointService.assignGuideUrl(this, "sharepointGuideUrl");
        }

        activateSharepoint () {
            return this.sharepointService.setSharepointUrl(this.exchangeId, this.sharepointUrl + this.sharepointUrlSuffix)
                .then(() => {
                    this.alerter.success(this.$scope.tr("sharepoint_set_url_success_message_text", this.exchangeId), this.$scope.alerts.dashboard);
                    this.$timeout(() => {
                        this.$location.path(`/configuration/sharepoint/${this.exchangeId}/${this.sharepointDomain}`);
                    }, 3000);
                })
                .catch((failure) => {
                    this.alerter.alertFromSWS(this.$scope.tr("sharepoint_set_url_failure_message_text"), failure, this.$scope.alerts.dashboard);
                });
        }
    });
