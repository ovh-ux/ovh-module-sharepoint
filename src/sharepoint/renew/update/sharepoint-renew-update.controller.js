angular
    .module("Module.sharepoint.controllers")
    .controller("SharepointUpdateRenewCtrl", class SharepointUpdateRenewCtrl {

        constructor (Alerter, $location, MicrosoftSharepointLicenseService, $q, $scope, $stateParams, $timeout) {
            this.alerter = Alerter;
            this.$location = $location;
            this.sharepointService = MicrosoftSharepointLicenseService;
            this.$q = $q;
            this.$scope = $scope;
            this.$stateParams = $stateParams;
            this.$timeout = $timeout;
        }

        $onInit () {
            this.timeout = null;

            this.exchangeId = this.$stateParams.exchangeId;
            this.search = { value: null };
            this.loaders = {
                init: true
            };
            this.buffer = {
                hasChanged: false,
                changes: []
            };

            this.getAccountIds();

            this.$scope.submit = () => {
                this.$q.all(
                    _.map(this.buffer.changes, (sharepoint) => this.sharepointService.updateSharepointAccount({
                        serviceName: this.exchangeId,
                        userPrincipalName: sharepoint.userPrincipalName,
                        deleteAtExpiration: sharepoint.deleteAtExpiration
                    }))
                )
                    .then(() => this.alerter.success(this.$scope.tr("exchange_update_billing_periode_success"), this.$scope.alerts.dashboard))
                    .catch((err) => this.alerter.alertFromSWS(this.$scope.tr("exchange_update_billing_periode_failure"), err, this.$scope.alerts.dashboard))
                    .finally(() => this.$scope.reset());

            };

            this.$scope.reset = () => {
                this.$location.search("action", null);
                this.$scope.resetAction();
            };
        }

        goSearch () {
            return this.getAccountIds();
        }

        emptySearch () {
            this.search.value = "";
            return this.getAccountIds();
        }

        getAccountIds () {
            this.loaders.init = true;
            this.accountIds = null;
            this.bufferedAccounts = [];

            this.sharepointService.getAccounts({
                serviceName: this.exchangeId,
                userPrincipalName: this.search.value
            }).then((accountIds) => {
                this.accountIds = accountIds;
            }).catch((err) => {
                this.alerter.alertFromSWS(this.$scope.tr("sharepoint_accounts_err"), err, this.$scope.alerts.dashboard);
            }).finally(() => {
                if (_.isEmpty(this.accountIds)) {
                    this.loaders.init = false;
                }
            });
        }

        onTranformItem (userPrincipalName) {
            return this.sharepointService.getAccountSharepoint({
                serviceName: this.exchangeId,
                userPrincipalName
            })
                .then((sharepoint) => {
                    sharepoint.userPrincipalName = userPrincipalName;
                    sharepoint.activated = true;
                    this.bufferedAccounts.push(_.clone(sharepoint));
                    return sharepoint;
                })
                .catch(() => ({
                    userPrincipalName,
                    activated: false
                }));
        }

        onTranformItemDone () {
            this.loaders.init = false;
        }

        changeRenew (account, newValue) {
            const buffered = _.find(this.bufferedAccounts, { userPrincipalName: account.userPrincipalName });

            if (buffered && buffered.deleteAtExpiration !== newValue) {
                this.buffer.changes.push(account);
            } else {
                _.remove(this.buffer.changes, account);
            }
            this.buffer.hasChanged = !!this.buffer.changes.length;
        }

        checkBuffer () {
            if (!this.accountIdsResume) {
                this.accountIdsResume = _.clone(this.accountIds, true);
            } else {
                this.accountIdsResume = [];
                _.delay(() => {
                    this.accountIdsResume = _.clone(this.accountIds, true);
                }, 50);
            }
        }

        onTranformItemResume (userPrincipalName) {
            return this.sharepointService.getAccountSharepoint({
                serviceName: this.exchangeId,
                userPrincipalName
            })
                .then((sharepoint) => {
                    sharepoint.userPrincipalName = userPrincipalName;
                    const buffered = _.find(this.buffer.changes, { userPrincipalName: sharepoint.userPrincipalName });

                    if (buffered) {
                        sharepoint.deleteAtExpiration = buffered.deleteAtExpiration;
                    }
                    return sharepoint;
                })
                .catch(() => null);
        }
    });
