angular
    .module("Module.sharepoint.controllers")
    .controller("SharepointTasksCtrl", class SharepointTasksCtrl {

        constructor ($scope, $stateParams, Alerter, MicrosoftSharepointLicenseService) {
            this.$scope = $scope;
            this.$stateParams = $stateParams;
            this.alerter = Alerter;
            this.sharepointService = MicrosoftSharepointLicenseService;
        }

        $onInit () {
            this.$scope.taskDetails = [];
            this.hasResult = false;
            this.loaders = {
                tasks: true,
                pager: false
            };

            this.getTasks();
        }

        getTasks () {
            this.loaders.tasks = true;
            this.tasksIds = null;

            return this.sharepointService.getTasks(this.$stateParams.exchangeId)
                .then((ids) => {
                    this.tasksIds = ids;
                    if (_.isArray(ids) && !_.isEmpty(ids)) {
                        this.hasResult = true;
                    }
                })
                .catch((err) => {
                    _.set(err, "type", err.type || "ERROR");
                    this.alerter.alertFromSWS(this.$scope.tr("sharepoint_tabs_tasks_error"), err, this.$scope.alerts.main);
                })
                .finally(() => {
                    if (_.isEmpty(this.tasksIds)) {
                        this.hasResult = false;
                        this.loaders.tasks = false;
                    }
                });
        }

        onTransformItem (taskId) {
            return this.sharepointService.getTask(this.$stateParams.exchangeId, taskId);
        }

        onTransformItemDone () {
            this.loaders.tasks = false;
            this.loaders.pager = false;
        }
    });
