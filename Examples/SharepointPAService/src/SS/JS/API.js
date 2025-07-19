/*
 * Create a new GUID
 * @returns {string} A newly generated UUID
 */
function generateUUID() {
  return "10000000-1000-4000-8000-100000000000".replace(/[018]/g, (c) =>
    (
      c ^
      (crypto.getRandomValues(new Uint8Array(1))[0] & (15 >> (c / 4)))
    ).toString(16)
  );
}

/*
 * Split an array into chunks of a specified size
 * @param {any[]} array - the array to split into chunks
 * @param {number} chunkSize - the chunk size desired
 * @returns {any[][]} - Array of chunks
 */
function splitArrayIntoChunks(array, chunkSize = 250) {
  const result = [];
  for (let i = 0; i < array.length; i += chunkSize) {
    result.push(array.slice(i, i + chunkSize));
  }
  return result;
}

class SharePointList {
  /**
   * @constructor
   * @example new SharePointList("https://severntrent.sharepoint.com/sites/OPERATIONALRISKMANAGEMENT/STORM", "STORM Missing Data")
   */
  constructor(site, listName) {
    this.site = site;
    this.list = listName;
    this.digest = {};
  }

  /**
   * Obtain list item type. Use `getListItemType` instead for cached value.
   * @protected
   * @example TBC
   */
  async getListItemTypeEx() {
    const f = await fetch(
      `${this.site}/_api/web/lists/GetByTitle('${this.list}')?$select=ListItemEntityTypeFullName`,
      { method: "GET", headers: { Accept: "application/json;odata=verbose" } }
    );
    let data = await f.json();
    return data.d.ListItemEntityTypeFullName;
  }

  /**
   * Obtain list item type, used in creation of items.
   * @example TBC
   */
  async getListItemType() {
    if (!this.itemType) this.itemType = await this.getListItemTypeEx();
    return this.itemType;
  }

  /**
   * Forcefully obtain request digest.
   * @remark Request digest will timeout after 1800 seconds. Use `getRequestDigest` instead to internally track this.
   * @example await myList.getRequestDigest()
   */
  async getRequestDigestEx() {
    const f = await fetch(`${this.site}/_api/contextinfo`, {
      method: "POST",
      headers: { Accept: "application/json;odata=verbose" },
    });
    let data = await f.json();
    return data.d.GetContextWebInformation.FormDigestValue;
  }

  /**
   * Obtain request digest.
   * @example await myList.getRequestDigest()
   */
  async getRequestDigest() {
    if (!this.digest.value || (Date.now() - this.digest.date) / 1000 > 1700) {
      this.digest = {
        value: await this.getRequestDigestEx(),
        date: Date.now(),
      };
    }
    return this.digest.value;
  }

  /*
   * Get a list items
   * @param {number} itemId - Sharepoint Item ID of item to get
   * @example await storm.getItem(1)
   * @docs https://learn.microsoft.com/en-us/sharepoint/dev/sp-add-ins/working-with-lists-and-list-items-with-rest
   */
  async getItem(itemId) {
    var f = await fetch(
      `${this.site}/_api/web/lists/GetByTitle('${this.list}')/items(${itemId})`,
      { method: "GET", headers: { Accept: "application/json;odata=verbose" } }
    );
    return (await f.json()).d;
  }

  /*
   * Update a list item
   * @param {string} itemID - Sharepoint Item ID of item to update
   * @param {object} data   - FieldName: Value list to update
   * @remark Including values of 'undefined' will automatically be excluded from 'JSON.stringify'.
   * @example await storm.setItem(1, {"County":"Staffordshire"})
   */
  async setItem(itemID, data) {
    let digest = await this.getRequestDigest();
    let itemType = await this.getListItemType();
    data["__metadata"] = { type: itemType };
    console.log({ caller: "setItem", itemID, digest, data });
    const f = await fetch(
      `${this.site}/_api/web/lists/GetByTitle('${this.list}')/items(${itemID})`,
      {
        method: "POST",
        headers: {
          Accept: "application/json;odata=verbose",
          "Content-Type": "application/json;odata=verbose",
          "Content-Length": JSON.stringify(data).length,
          "If-Match": "*",
          "X-HTTP-Method": "MERGE",
          "X-RequestDigest": digest,
        },
        body: JSON.stringify(data),
      }
    );
    return f;
  }

  /*
   * Add a list item
   * @param {object} data   - Data to add
   * @example await storm.addItem({"County":"Staffordshire", ...})
   */
  async addItem(data, itemType = undefined) {
    let digest = await this.getRequestDigest();
    if (!itemType) itemType = await this.getListItemType();
    data["__metadata"] = { type: itemType };
    console.log({ caller: "addItem", itemType, digest, data });
    const f = await fetch(
      `${this.site}/_api/web/lists/GetByTitle('${this.list}')/items`,
      {
        method: "POST",
        headers: {
          Accept: "application/json;odata=verbose",
          "Content-Type": "application/json;odata=verbose",
          "Content-Length": JSON.stringify(data).length,
          "X-RequestDigest": digest,
        },
        body: JSON.stringify(data),
      }
    );
    return f;
  }

  /*
   * Get all list items
   * @example await myList.getListItems()
   * @docs https://learn.microsoft.com/en-us/sharepoint/dev/sp-add-ins/working-with-lists-and-list-items-with-rest
   */
  async getItems() {
    var f = await fetch(
      `${this.site}/_api/web/lists/GetByTitle('${this.list}')/items`,
      { method: "GET", headers: { Accept: "application/json;odata=verbose" } }
    );
    var data = (await f.json()).d;
    var results = data.results;
    while (!!data.__next) {
      f = await fetch(data.__next, {
        method: "GET",
        headers: { Accept: "application/json;odata=verbose" },
      });
      data = (await f.json()).d;
      results = results.concat(data.results);
    }
    return results;
  }

  /*
   * Get list items filtering to an ODATA Query
   * @example await myList.getItemsWithODataQuery("myCol eq 'test'")
   * @remark ODATA Queries do not work effectively on lists with more than 500 records. It's recommended to use `getItemsWithInplaceQuery` instead where possible.
   * @docs https://learn.microsoft.com/en-us/sharepoint/dev/sp-add-ins/working-with-lists-and-list-items-with-rest
   */
  async getItemsWithODataQuery(oDataQuery) {
    var f = await fetch(
      `${this.site}/_api/web/lists/GetByTitle('${this.list}')/items?$filter=${oDataQuery}`,
      { method: "GET", headers: { Accept: "application/json;odata=verbose" } }
    );
    var data = (await f.json()).d;
    var results = data.results;
    while (!!data.__next) {
      f = await fetch(data.__next, {
        method: "GET",
        headers: { Accept: "application/json;odata=verbose" },
      });
      data = (await f.json()).d;
      results = results.concat(data.results);
    }
    return results;
  }

  /*
   * Get items using an inplace Query
   * @example
   */
  async getItemsWithInplaceQuery(inplaceQuery) {
    let data = await fetch(
      `${this.site}/_api/web/lists/GetByTitle('${this.list}')/RenderListDataAsStream?InplaceSearchQuery=${inplaceQuery}`,
      {
        method: "POST",
        headers: { "content-type": "application/json;odata=verbose" },
      }
    );
    let json = await data.json();
    return json["Row"];
  }

  /*
   * Update multiple items in a batch query
   * @param {IChange[]}          data        - List of items to change
   * @param {string|Null = null} changeSetId - UUID of the change set. By default a random UUID is generated.
   * @param {string|Null = null} batchUuid   - UUID of the batch. By default a random UUID is generated.
   * @type IChange = {id: number, data: object} - data = object of  fields to change and values to change them to.
   * @remark https://sharepoint.stackexchange.com/questions/234766/batch-update-create-list-items-using-rest-api-in-sharepoint-2013
   * @remark https://learn.microsoft.com/en-us/sharepoint/dev/sp-add-ins/make-batch-requests-with-the-rest-apis
   * @example ```
   *   await db.setItemsBatch([
   *     {id: 8369, data:{NonInfraOperationalArea: "N/A"}},
   *     {id: 8370, data:{NonInfraOperationalArea: "Stafford"}},
   *     {id: 8371, data:{NonInfraOperationalArea: "Stratford and Warwick"}},
   *   ])
   * ```
   */
  async setItemsBatch(batch, changeSetId = null, batchUuid = null) {
    let itemType = await this.getListItemType();
    var changeSetId, batchUuid;
    if (changeSetId == null) changeSetId = generateUUID();
    if (batchUuid == null) batchUuid = generateUUID();
    let responses = [];

    //A max of 1000 operations are allowed in a changeset; To stay well under this value, we use batches of 750.
    for (let data of splitArrayIntoChunks(batch, 750)) {
      //Workaround for issue described @ https://learn.microsoft.com/en-us/answers/questions/1383519/error-while-trying-to-bulk-update-sharepoint-onlin
      data.push(data[0]);

      //Create changeset
      let batch = [];
      batch.push(`--batch_${batchUuid}`);
      batch.push(
        `Content-Type: multipart/mixed; boundary=changeset_${changeSetId}`
      );
      batch.push("");

      data.forEach((item) => {
        item.data["__metadata"] = { type: itemType }; //Bind ItemType

        //Add change
        batch.push(`--changeset_${changeSetId}`);
        batch.push("Content-Type:application/http");
        batch.push("Content-Transfer-Encoding: binary");
        batch.push("");
        batch.push(
          `PATCH ${this.site}/_api/web/lists/GetByTitle('${this.list}')/items(${item.id}) HTTP/1.1`
        );
        batch.push(`Content-Type: application/json;odata=verbose;`);
        batch.push("Accept: application/json");
        batch.push("If-Match: *");
        batch.push("X-HTTP-Method: MERGE");
        batch.push("");
        batch.push(JSON.stringify(item.data));
        batch.push("");
      });

      //End changeset to create Data
      batch.push(`--changeSetId_${changeSetId}--`);

      console.log({
        caller: "setItemsBatch",
        data,
        batch,
        batchUuid,
        changeSetId,
      });

      const f = await fetch(`${this.site}/_api/$batch`, {
        method: "POST",
        headers: {
          "X-RequestDigest": await this.getRequestDigest(),
          "Content-Type": `multipart/mixed; boundary="batch_${batchUuid}"`,
        },
        body: batch.join("\r\n"),
      });
      responses.push(f);
    }
    return responses;
  }
}
