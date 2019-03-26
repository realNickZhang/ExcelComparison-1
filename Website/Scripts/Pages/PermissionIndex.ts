﻿/// <reference path="../../wwwroot/lib/@types/jquery/index.d.ts" />
/// <reference path="../Models/ClientPageInfo.ts" />
/// <reference path="../Models/ActiveState.ts" />
/// <reference path="../Shared/ControlHelper.ts" />
/// <reference path="../Shared/ClientPageInfoHelper.ts" />

class PermissionIndex {
	private _dataUrl: string = '/permission/readall'
	private _pageSize: number = 10;
	private _activeState: ActiveState;

	constructor() {
		this._activeState = new ActiveState();
		this.setupGrid();
	}

	private setupGrid() {
		var self = this;
		var grid: any = $("#jsGrid");
		var tokenelement = $('input[name="__RequestVerificationToken"]');
		var token = tokenelement.val();
		grid.jsGrid({
			width: "100%",
			height: "600px",

			inserting: false,
			editing: false,
			sorting: true,
			paging: true,

			datatype: 'json',
			autoload: true,
			pageLoading: true,
			pageSize: self._pageSize,
			controller: {
				loadData: function (filter: any) {
					var deferred = $.Deferred();
					if (typeof filter.sortField == 'undefined') {
						filter.sortField = 'id';
						filter.sortOrder = 'asc';
					}
					var postData: ClientPageInfo = ClientPageInfoHelper.GetClientPageInfo(filter, self._pageSize, self._activeState);
					$.ajax({
						type: 'POST',
						url: self._dataUrl,
						data: {
							__RequestVerificationToken: token,
							clientPageInfo: postData
						},
						dataType: 'json'
					}).done(function (response: any) {
						deferred.resolve({ data: response.Data, itemsCount: response.Total });
					});

					return deferred.promise();
				}
			},

			fields: [
				{ name: "Name", title: "Permission Name", type: "text", width: 150, validate: "required" },
				{ name: "Action", title: "Action", type: "text", width: 200 },
				{
					type: "control",
					width: 50,
					itemTemplate: function (value: any, item: any) {
						var $iconPencil = $("<i>").attr({ class: "far fa-edit" });
						var $iconTrash = $("<i>").attr({ class: "far fa-trash-alt" });

						var $customEditButton = $("<a>")
							.attr({ class: "customButton btn btn-primary btn-xs" })
							.attr({ href: "/Permission/Update/" + item.Id })
							.append($("<span>").append($iconPencil));

						var $customDeleteButton = $("<a>")
							.attr({ class: "customButton btn btn-danger btn-xs" })
							.attr({ href: "/Permission/Delete/" + item.Id })
							.append($("<span>").append($iconTrash));

						return $("<div>").append($customEditButton).append($customDeleteButton);
					}
				}
			]
		});
	}

	public changeActiveState(value: string) {
		this._activeState = ControlHelper.ChangeActiveState(value);
		ControlHelper.ReloadGrid();
	}
}