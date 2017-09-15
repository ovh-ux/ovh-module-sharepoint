{
    const URL_DEFAULT_SUFFIX = ".sp.ovh.net";
    const DEFAULT_SUBSIDIARY = "GB";

    angular
        .module("Module.sharepoint.services")
        .service("MicrosoftSharepointLicenseService", class MicrosoftSharepointLicenseService {

            constructor (Alerter, OvhHttp, Products, $q, SHAREPOINT_GUIDE_URLS, translator, User) {
                this.Alerter = Alerter;
                this.OvhHttp = OvhHttp;
                this.Products = Products;
                this.$q = $q;
                this.SHAREPOINT_GUIDE_URLS = SHAREPOINT_GUIDE_URLS;
                this.translator = translator;
                this.User = User;
            }

            $onInit () {
                this.cache = {
                    models: "UNIVERS_MODULE_SHAREPOINT_MODELS",
                    sharepoints: "UNIVERS_MODULE_SHAREPOINT_SHAREPOINTS",
                    services: "UNIVERS_MODULE_SHAREPOINT_SERVICES",
                    servicesInfos: "UNIVERS_MODULE_SHAREPOINT_SERVICE_INFOS",
                    accounts: "UNIVERS_MODULE_SHAREPOINT_SERVICE_ACCOUNTS",
                    license: "UNIVERS_MODULE_SHAREPOINT_SERVICE_LICENSE"
                };

                this.User
                    .getUrlOfEndsWithSubsidiary("express_order")
                    .then((orderBaseUrl) => {
                        this.orderBaseUrl = orderBaseUrl;
                    })
                    .catch((error) => {
                        this.Alerter.alertFromSWS(this.translator.tr("sharepoint_dashboard_error"), error);
                    });
            }

            /**
             * Set guide
             * @param  {string} assignToObject
             * @param  {string} assignToProperty
             */
            assignGuideUrl (assignToObject, assignToProperty) {
                return this.User
                    .getUser()
                    .then((user) => {
                        assignToObject[assignToProperty] = this.SHAREPOINT_GUIDE_URLS[user.ovhSubsidiary] || this.SHAREPOINT_GUIDE_URLS[DEFAULT_SUBSIDIARY];
                    });
            }

            /**
             * Get serviceName infos
             * @param  {string} exchangeId
             */
            retrievingMSService (exchangeId) {
                return this.OvhHttp
                    .get(`/msServices/${exchangeId}`, {
                        rootPath: "apiv6",
                        cache: this.cache.sharepoints
                    });
            }

            /**
             * Get sharepoint infos
             * @param  {string} exchangeId
             */
            getSharepoint (exchangeId) {
                return this.OvhHttp
                    .get(`/msServices/${exchangeId}/sharepoint`, {
                        rootPath: "apiv6",
                        cache: this.cache.sharepoints
                    });
            }

            /**
             * Update sharepoint
             * @param {string} exchangeId
             * @param {string} url
             */
            setSharepointUrl (exchangeId, url) {
                return this.OvhHttp
                    .put(`/msServices/${exchangeId}/sharepoint`, {
                        rootPath: "apiv6",
                        data: {
                            url
                        },
                        clearAllCache: this.cache.sharepoints
                    });
            }

            /**
             * @param  {string} exchangeName
             * @param  {string[]]} emails
             */
            getSharepointOrderUrl (exchangeName, emails) {
                if (_.isEmpty(this.orderBaseUrl)) {
                    return null;
                }

                const configuration = emails.map((email) => ({
                    planCode: "sharepoint_account",
                    configuration: [{
                        label: "EXCHANGE_ACCOUNT_ID",
                        values: [email]
                    }]
                }));

                const products = [{
                    planCode: "sharepoint_platform",
                    configuration: [{
                        label: "EXCHANGE_SERVICE_NAME",
                        values: [
                            exchangeName
                        ]
                    }],
                    option: configuration
                }];

                return `${this.orderBaseUrl}#/new/express/resume?productId=sharepoint&products=${JSURL.stringify(products)}`;
            }

            /**
             *
             * @param  {string} serviceName
             * @param  {string} primaryEmailAddress
             */
            getSharepointAccountOrderUrl (serviceName, primaryEmailAddress) {
                if (_.isEmpty(this.orderBaseUrl)) {
                    return null;
                }

                const productId = "sharepoint";
                const products = [{
                    planCode: "sharepoint_account",
                    configuration: [{
                        label: "EXCHANGE_ACCOUNT_ID",
                        values: [
                            primaryEmailAddress
                        ]
                    }]
                }];

                return `${this.orderBaseUrl}#/new/express/resume?productId=${productId}&serviceName=${serviceName}&products=${JSURL.stringify(products)}`;
            }

            /**
             *
             * @param  {string} serviceName
             * @param  {number} number
             */
            getSharepointStandaloneNewAccountOrderUrl (serviceName, number) {
                if (_.isEmpty(this.orderBaseUrl)) {
                    return null;
                }

                const products = [{
                    planCode: "sharepoint_account",
                    quantity: number || 1,
                    productId: "sharepoint",
                    serviceName
                }];

                return `${this.orderBaseUrl}#/new/express/resume?products=${JSURL.stringify(products)}`;
            }

            /**
             *
             * @param  {number} quantity
             */
            getSharepointStandaloneOrderUrl (quantity) {
                if (_.isEmpty(this.orderBaseUrl)) {
                    return null;
                }

                const productId = "sharepoint";
                const products = [{
                    planCode: "sharepoint_platform",
                    configuration: [],
                    option: [{
                        planCode: "sharepoint_account",
                        quantity: quantity || 1,
                        configuration: []
                    }]
                }];

                return `${this.orderBaseUrl}#/new/express/resume?productId=${productId}&products=${JSURL.stringify(products)}`;
            }

            /**
             * Get sharepoint options
             * @param  {string} serviceName
             */
            retrievingSharepointServiceOptions (serviceName) {
                return this.OvhHttp
                    .get(`/order/cartServiceOption/sharepoint/${serviceName}`, {
                        rootPath: "apiv6",
                        cache: this.cache.sharepoints
                    });
            }

            /**
             *
             * @param  {Object} opts
             */
            getAccounts (opts) {
                const queryParam = {};

                if (_.isEmpty(opts.userPrincipalName)) {
                    queryParam.userPrincipalName = `%${opts.userPrincipalName}%`;
                }

                return this.OvhHttp.get(`/msServices/${opts.serviceName}/account`, {
                    rootPath: "apiv6",
                    params: queryParam,
                    cache: this.cache.sharepoints
                });
            }

            /**
             *
             * @param  {Object} opts
             */
            restoreAdminRights (opts) {
                return this.OvhHttp
                    .post(`/msServices/${opts.serviceName}/sharepoint/restoreAdminRights`, {
                        rootPath: "apiv6"
                    });
            }

            /**
             *
             * @param  {Object} opts
             */
            getAccountDetails (opts) {
                return this.OvhHttp
                    .get(`/msServices/${opts.serviceName}/account/${opts.userPrincipalName}`, {
                        rootPath: "apiv6",
                        cache: this.cache.sharepoints
                    });
            }

            /**
             *
             * @param  {Object} opts
             */
            getAccountSharepoint (opts) {
                return this.OvhHttp
                    .get(`/msServices/${opts.serviceName}/account/${opts.userPrincipalName}/sharepoint`, {
                        rootPath: "apiv6",
                        cache: this.cache.sharepoints
                    });
            }

            /**
             *
             * @param  {Object} data
             */
            updateSharepointAccount (data) {
                return this.OvhHttp
                    .put(`/msServices/${data.serviceName}/account/${data.userPrincipalName}/sharepoint`, {
                        rootPath: "apiv6",
                        data,
                        clearAllCache: this.cache.sharepoints
                    });
            }

            /**
             *
             * @param  {string} exchangeId
             * @param  {string} userPrincipalName
             * @param  {Object} data
             */
            updateSharepoint (exchangeId, userPrincipalName, data) {
                return this.OvhHttp
                    .put(`/msServices/${exchangeId}/account/${userPrincipalName}`, {
                        rootPath: "apiv6",
                        data,
                        clearAllCache: this.cache.sharepoints
                    });
            }

            /**
             *
             * @param  {Object} opts
             */
            updatingSharepointPasswordAccount (opts) {
                return this.OvhHttp
                    .post(`/msServices/${opts.serviceName}/account/${opts.userPrincipalName}/changePassword`, {
                        rootPath: "apiv6",
                        data: {
                            password: opts.password
                        }
                    });
            }

            /**
             *
             * @param  {string} exchangeId
             * @param  {string} account
             */
            deleteSharepointAccount (exchangeId, account) {
                return this.OvhHttp
                    .put(`/msServices/${exchangeId}/account/${account}/sharepoint`, {
                        rootPath: "apiv6",
                        data: {
                            deleteAtExpiration: true
                        },
                        clearAllCache: this.cache.sharepoints
                    });
            }

            /**
             * @param  {Object} opts
             */
            getAccount (opts) {
                return this.OvhHttp
                    .get(`/sharepoint/${opts.organizationName}/service/${opts.sharepointService}/account/${opts.userPrincipalName}`, {
                        rootPath: "apiv6",
                        cache: this.cache.accounts
                    });
            }

            getAccountTasks (opts) {
                return this.OvhHttp
                    .get(`/sharepoint/${opts.organizationName}/service/${opts.sharepointService}/account/${opts.userPrincipalName}/tasks`, {
                        rootPath: "apiv6"
                    });
            }

            /**
             * @param  {string} exchangeId
             */
            retrievingExchangeOrganization (exchangeId) {
                return this.Products
                    .getProductsByType()
                    .then((productsByType) => {
                        const exchange = _.find(productsByType.exchanges, { name: exchangeId });

                        // If Sharepoint standalone, no exchange service attached to it.
                        return exchange ? exchange.organization : null;
                    });
            }

            /**
             * An API function will be developed to get directly the info.
             * For now, an exchange with hostname "ex.mail.ovh.net" should have the suffix ".sp.ovh.net"
             * An exchange with hostname "ex2.mail.ovh.net" will have the suffix ".sp2.ovh.net"
             */
            retrievingSharepointSuffix (serviceName) {
                return this.OvhHttp
                    .get("/msServices/{serviceName}/sharepoint", {
                        rootPath: "apiv6",
                        urlParams: {
                            serviceName
                        }
                    })
                    .then((sharepoint) => {
                        const separator = _.startsWith(sharepoint.farmUrl, ".") ? "" : ".";

                        return `${separator}${sharepoint.farmUrl}`;
                    })
                    .catch(() => URL_DEFAULT_SUFFIX);
            }

            /**
             * Get upn suffixes, that is the domains allowed for account's configuration
             */
            getUsedUpnSuffixes () {
                return this.OvhHttp
                    .get("/msServices", {
                        rootPath: "apiv6",
                        cache: this.cache.sharepoints
                    })
                    .then((msServices) => {
                        const queue = msServices.map((serviceId) => this.OvhHttp
                            .get(`/msServices/${serviceId}/upnSuffix`, {
                                rootPath: "apiv6",
                                cache: this.cache.sharepoints
                            })
                            .then((suffixes) => suffixes)
                            .catch(() => null));

                        return this.$q.all(queue).then((data) => _.flatten(_.compact(data)));
                    })
                    .catch(() => []);
            }

            /**
             * Get upn suffixes, that is the domains allowed for account's configuration
             */
            getSharepointUpnSuffixes (exchangeId) {
                return this.OvhHttp
                    .get(`/msServices/${exchangeId}/upnSuffix`, {
                        rootPath: "apiv6",
                        cache: this.cache.sharepoints
                    });
            }

            /**
             * Add an upn suffix
             */
            addSharepointUpnSuffixe (exchangeId, suffix) {
                return this.OvhHttp
                    .post(`/msServices/${exchangeId}/upnSuffix`, {
                        rootPath: "apiv6",
                        data: {
                            suffix
                        },
                        clearAllCache: this.cache.sharepoints
                    });
            }

            /**
             * Delete an upn suffix
             */
            deleteSharepointUpnSuffix (exchangeId, suffix) {
                return this.OvhHttp
                    .delete(`/msServices/${exchangeId}/upnSuffix/${suffix}`, {
                        rootPath: "apiv6",
                        clearAllCache: this.cache.sharepoints
                    });
            }

            /**
             * Get upn suffix details
             */
            getSharepointUpnSuffixeDetails (exchangeId, suffix) {
                return this.OvhHttp
                    .get(`/msServices/${exchangeId}/upnSuffix/${suffix}`, {
                        rootPath: "apiv6",
                        cache: this.cache.sharepoints
                    });
            }
        });
}
