angular
    .module("Module.sharepoint.controllers")
    .controller("SharepointUpdateAccountCtrl", class SharepointUpdateAccountCtrl {

        constructor ($scope, $q, $stateParams, Alerter, MicrosoftSharepointLicenseService) {
            this.$scope = $scope;
            this.$q = $q;
            this.$stateParams = $stateParams;
            this.Alerter = Alerter;
            this.SharepointService = MicrosoftSharepointLicenseService;
        }

        $onInit () {
            this.exchangeId = this.$stateParams.exchangeId;

            this.originalValue = angular.copy(this.$scope.currentActionData);

            this.account = this.$scope.currentActionData;
            this.account.login = this.account.userPrincipalName.split("@")[0];
            this.account.domain = this.account.userPrincipalName.split("@")[1];

            this.availableDomains = [];
            this.availableDomains.push(this.account.domain);

            this.$scope.updateMsAccount = () => this.updateMsAccount();

            this.getMsService();
            this.getAccountDetails();
            this.getSharepointUpnSuffixes();
        }

        getMsService () {
            this.SharepointService.retrievingMSService(this.exchangeId)
                .then((exchange) => { this.$scope.exchange = exchange; });
        }

        getAccountDetails () {
            this.SharepointService.getAccountDetails(this.exchangeId, this.account.userPrincipalName)
                .then((accountDetails) => _.assign(this.account, accountDetails));
        }

        getSharepointUpnSuffixes () {
            this.SharepointService.getSharepointUpnSuffixes(this.exchangeId)
                .then((upnSuffixes) =>

                    // check and filter the domains if they are not validated
                    this.$q.all(
                        _.filter(upnSuffixes, (suffix) => this.SharepointService.getSharepointUpnSuffixeDetails(this.exchangeId, suffix)
                            .then((suffixDetails) => suffixDetails.ownershipValidated))
                    )
                )
                .then((availableDomains) => { this.availableDomains = _.union([this.account.domain], availableDomains); });
        }

        updateMsAccount () {
            this.account.userPrincipalName = `${this.account.login}@${this.account.domain}`;
            return this.SharepointService.updateSharepoint(this.exchangeId, this.originalValue.userPrincipalName, _.pick(this.account, ["userPrincipalName", "firstName", "lastName", "initials", "displayName"]))
                .then(() => {
                    this.Alerter.success(this.$scope.tr("sharepoint_account_update_configuration_confirm_message_text", this.account.userPrincipalName), this.$scope.alerts.main);
                })
                .catch((err) => {
                    this.Alerter.alertFromSWS(this.$scope.tr("sharepoint_account_update_configuration_error_message_text"), err, this.$scope.alerts.main);
                })
                .finally(() => {
                    this.$scope.resetAction();
                });
        }
    });
