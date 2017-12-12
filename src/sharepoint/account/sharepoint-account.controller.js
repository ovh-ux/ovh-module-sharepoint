angular
    .module("Module.sharepoint.controllers")
    .controller("SharepointAccountsCtrl", class SharepointAccountsCtrl {

        constructor ($scope, $location, $stateParams, $timeout, Alerter, MicrosoftSharepointLicenseService, Poller) {
            this.$scope = $scope;
            this.$location = $location;
            this.$stateParams = $stateParams;
            this.$timeout = $timeout;
            this.alerter = Alerter;
            this.sharepointService = MicrosoftSharepointLicenseService;
            this.pollerService = Poller;
        }

        $onInit () {
            this.timeout = null;
            this.search = {};
            this.accounts = [];
            this.exchangeId = this.$stateParams.exchangeId;
            this.hasResult = false;
            this.loaders = {
                init: true,
                search: false
            };
            this.iteration = 0;

            this.getSharepoint();
            this.getExchangeOrganization();
            this.getAccountIds();

            this.$scope.$on("$destroy", () => {
                this.pollerService.kill({
                    namespace: "sharepoint.accounts.poll"
                });
            });
        }

        updateSearch () {
            return this.getAccountIds();
        }

        emptySearch () {
            this.search.value = "";
            this.updateSearch();
        }

        getSharepoint () {
            this.sharepointService.getSharepoint(this.$stateParams.exchangeId)
                .then((sharepoint) => {
                    this.sharepoint = sharepoint;
                    if (_.isNull(this.sharepoint.url)) {
                        this.$location.path(`/configuration/sharepoint/${this.$stateParams.exchangeId}/${this.sharepoint.domain}/setUrl`);
                    }
                })
                .catch((err) => {
                    _.set(err, "type", err.type || "ERROR");
                    this.alerter.alertFromSWS(this.$scope.tr("sharepoint_dashboard_error"), err, this.$scope.alerts.main);
                });
        }

        getExchangeOrganization () {
            this.sharepointService.retrievingExchangeOrganization(this.$stateParams.exchangeId)
                .then((organization) => {
                    this.isStandAlone = _.isNull(organization);
                });
        }

        updateSharepoint (account, type, officeLicense) {
            return this.sharepointService.updateSharepointAccount(this.exchangeId, account.userPrincipalName, {
                accessRights: type,
                officeLicense
            }).then(() => {
                this.search.value = "";
                this.hasResult = false;
                this.getAccountIds();
            });
        }

        startPoller (userPrincipalName) {
            this.pollerService.poll(`apiv6/msServices/${this.exchangeId}/account/${userPrincipalName}/`, null, {
                interval: 15000,
                successRule: { state: (account) => account.taskPendingId === 0 },
                namespace: "sharepoint.accounts.poll"
            }).then((account) => {
                account.activated = true;
                account.userPrincipalName = userPrincipalName;
                const index = _.findIndex(this.accounts, { userPrincipalName });
                if (index > -1) {
                    this.accounts[index] = account;
                }
            }).catch(() => {
                this.pollerService.kill({
                    namespace: "sharepoint.accounts.poll"
                });
            });
        }

        activateSharepointUser (account) {
            if (!account.taskPendingId) {
                if (account.accessRights === "administrator") {
                    this.updateSharepoint(account, "user");
                }
            }
        }

        activateSharepointAdmin (account) {
            if (!account.taskPendingId) {
                if (account.accessRights === "user") {
                    this.updateSharepoint(account, "administrator");
                }
            }
        }

        activateSharepoint (account) {
            if (!account.taskPendingId) {
                window.open(this.sharepointService.getSharepointAccountOrderUrl(this.$stateParams.productId, account.userPrincipalName), "_blank");
            }
        }

        activateOfficeLicence (account) {
            if (!account.taskPendingId) {
                this.updateSharepoint(account, null, true);
            }
        }

        deactivateOfficeLicence (account) {
            if (!account.taskPendingId) {
                this.updateSharepoint(account, null, false);
            }
        }

        deactivateSharepoint (account) {
            if (!account.taskPendingId) {
                this.$scope.setAction("account/delete/sharepoint-account-delete", account);
            }
        }

        activateOffice (account) {
            if (!account.taskPendingId) {
                this.$scope.setAction("activate-office/sharepoint-activate", account);
            }
        }

        updatePassword (account) {
            if (!account.taskPendingId) {
                this.$scope.setAction("account/password/update/sharepoint-account-password-update", account);
            }
        }

        updateAccount (account) {
            if (!account.taskPendingId && this.isStandAlone) {
                this.$scope.setAction("account/update/sharepoint-account-update", account);
            }
        }

        getAccountIds () {
            this.loaders.search = true;
            this.accountIds = null;

            return this.sharepointService.getAccounts(this.exchangeId, this.search.value)
                .then((accountIds) => {
                    this.accountIds = accountIds;
                }).catch((err) => {
                    _.set(err, "type", err.type || "ERROR");
                    this.alerter.alertFromSWS(this.$scope.tr("sharepoint_accounts_err"), err, this.$scope.alerts.main);
                }).finally(() => {
                    if (_.isEmpty(this.accountIds)) {
                        this.loaders.search = false;
                    } else {
                        this.hasResult = true;
                    }
                    this.loaders.init = false;
                });
        }

        resetAdminRights () {
            this.$scope.setAction("admin-rights/reset/sharepoint-admin-rights-reset");
        }

        activateAcount () {
            this.$scope.setAction("activate-account/sharepoint-activate-account");
        }

        onTranformItem (userPrincipalName) {
            return this.sharepointService.getAccountSharepoint(this.exchangeId, userPrincipalName)
                .then((sharepoint) => {
                    sharepoint.userPrincipalName = userPrincipalName;
                    sharepoint.activated = true;
                    sharepoint.usedQuota = filesize(sharepoint.currentUsage, { standard: "iec", output: "object" });
                    sharepoint.totalQuota = filesize(sharepoint.quota, { standard: "iec", output: "object" });
                    sharepoint.percentUse = Math.round((sharepoint.currentUsage / sharepoint.quota) * 100);
                    if (sharepoint.taskPendingId > 0) {
                        this.startPoller(userPrincipalName);
                    }
                    return sharepoint;
                }).catch(() => ({
                    userPrincipalName,
                    activated: false
                }));
        }

        onTranformItemDone () {
            this.loaders.search = false;
        }
    });
