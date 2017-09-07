angular
    .module("Module.sharepoint.controllers")
    .controller("SharepointUpdateAccountCtrl", class SharepointUpdateAccountCtrl {

        constructor (Alerter, MicrosoftSharepointLicenseService, $q, $stateParams, $scope) {
            this.alerter = Alerter;
            this.sharepointService = MicrosoftSharepointLicenseService;
            this.$q = $q;
            this.$stateParams = $stateParams;
            this.$scope = $scope;
        }

        $onInit () {
            this.exchangeId = this.$stateParams.exchangeId;

            this.originalValue = angular.copy(this.$scope.currentActionData);

            this.account = this.$scope.currentActionData;
            this.account.login = this.account.userPrincipalName.split("@")[0];
            this.account.domain = this.account.userPrincipalName.split("@")[1];

            this.availableDomains = [];
            this.availableDomains.push(this.account.domain);

            this.getMsService();
            this.getAccountDetails();
            this.getSharepointUpnSuffixes();

            this.$scope.updateMsAccount = () => {
                this.account.userPrincipalName = `${this.account.login}@${this.account.domain}`;

                this.sharepointService.updateSharepoint(this.exchangeId,
                                                        this.originalValue.userPrincipalName,
                                                        _.pick(this.account, ["userPrincipalName", "firstName", "lastName", "initials", "displayName"])
                )
                    .then(() => this.alerter.success(this.$scope.tr("sharepoint_account_update_configuration_confirm_message_text", this.account.userPrincipalName), this.$scope.alerts.dashboard))
                    .catch((err) => this.alerter.alertFromSWS(this.$scope.tr("sharepoint_account_update_configuration_error_message_text"), err, this.$scope.alerts.dashboard))
                    .finally(() => this.$scope.resetAction());
            };
        }

        getMsService () {
            this.sharepointService.retrievingMSService(this.exchangeId)
                .then((exchange) => { this.$scope.exchange = exchange; });
        }

        getAccountDetails () {
            this.sharepointService.getAccountDetails({ serviceName: this.exchangeId, userPrincipalName: this.account.userPrincipalName })
                .then((accountDetails) => _.assign(this.account, accountDetails));
        }

        getSharepointUpnSuffixes () {
            this.sharepointService.getSharepointUpnSuffixes(this.exchangeId)
                .then((upnSuffixes) =>

                    // check and filter the domains if they are not validated
                    this.$q.all(
                        _.filter(upnSuffixes, (suffix) => this.sharepointService.getSharepointUpnSuffixeDetails(this.exchangeId, suffix)
                            .then((suffixDetails) => suffixDetails.ownershipValidated))
                    )
                )
                .then((availableDomains) => { this.availableDomains = _.union([this.account.domain], availableDomains); });
        }
    });
