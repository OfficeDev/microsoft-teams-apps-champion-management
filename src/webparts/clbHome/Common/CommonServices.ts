import { WebPartContext } from "@microsoft/sp-webpart-base";
import { sp } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import siteconfig from "../config/siteconfig.json";
import * as stringsConstants from "../constants/strings";
import { IRectangle } from '@fluentui/react/lib/Utilities';
import { IConfigList } from "../components/ManageConfigSettings";

let rootSiteURL: string;
export default class CommonServices {

  constructor(context: WebPartContext, siteUrl: string) {

    //Set context for PNP   
    //When App is added to a Teams
    if (context.pageContext.web.serverRelativeUrl == "/")
      rootSiteURL = context.pageContext.web.absoluteUrl;
    //when app is added as personal app
    else
      rootSiteURL = siteUrl;

    //Set up URL for CMP site
    rootSiteURL = rootSiteURL + "/" + siteconfig.inclusionPath + "/" + siteconfig.sitename;
    sp.setup({
      sp: {
        baseUrl: rootSiteURL
      },
    });
  }

  //Method to get the pixel height for a given page
  public getPageHeight = (rowHeight: number, ROWS_PER_PAGE: number): number => {
    return rowHeight * ROWS_PER_PAGE;
  }

  //Method to get how many items to render per page from specified index
  public getItemCountForPage = (itemIndex: number, surfaceRect: IRectangle, MAX_ROW_HEIGHT: number, ROWS_PER_PAGE: number) => {
    let columnCount: number;
    let rowHeight: number;
    if (itemIndex === 0) {
      columnCount = Math.ceil(surfaceRect.width / MAX_ROW_HEIGHT);
      rowHeight = Math.floor(surfaceRect.width / columnCount);
    }
    return { itemCountForPage: columnCount * ROWS_PER_PAGE, columnCount: columnCount, rowHeight: rowHeight };
  }

  //Get Member list column Config settings
  public async getMemberListColumnConfigSettings() {
    const filterQuery = "Title eq '" + stringsConstants.RegionColumn + "' or Title eq '"
      + stringsConstants.CountryColumn + "' or Title eq '" + stringsConstants.GroupColumn + "'";
    const configListData: IConfigList[] = await this.getFilteredListItemsWithSpecificColumns(
      stringsConstants.ConfigList,
      `${stringsConstants.TitleColumn},${stringsConstants.ValueColumn},${stringsConstants.IDColumn}`,
      filterQuery
    );
    return configListData;
  }

  //Get Member List Column Display Names
  public async getMemberListColumnDisplayNames() {
    const columnsFilter = "InternalName eq '" + stringsConstants.RegionColumn + "' or InternalName eq '"
      + stringsConstants.CountryColumn + "' or InternalName eq '" + stringsConstants.GroupColumn + "'";
    const columnsDisplayNames: any[] = await this.getColumnsDisplayNames(stringsConstants.MemberList, columnsFilter);
    return columnsDisplayNames;
  }

  //Get list items based on only a filter
  public async getItemsWithOnlyFilter(listname: string, filterparametres: any): Promise<any> {
    var items: any[] = [];
    items = await sp.web.lists.getByTitle(listname).items.filter(filterparametres).getAll();
    return items;
  }

  //Get list items based on only a filter and sorted
  public async getItemsSortedWithFilter(listname: string, filterparametres: any, descColumn: any): Promise<any> {
    var items: any[] = [];
    items = await sp.web.lists.getByTitle(listname).items.filter(filterparametres).orderBy(descColumn, false)();
    return items;
  }
  //Get all items from a list
  public async getAllListItems(listname: string): Promise<any> {
    var items: any[] = [];
    items = await sp.web.lists.getByTitle(listname).items.getAll();
    return items;
  }
  //Get all list items with specific columns
  public async getAllItemsWithSpecificColumns(listname: string, columns: string): Promise<any> {
    var items: any[] = [];
    items = await sp.web.lists.getByTitle(listname).items.select(columns).getAll();
    return items;
  }

  //Get all items with paged from a list
  public async getAllListItemsPaged(listname: string): Promise<any> {
    var items: any = [];
    items = await sp.web.lists.getByTitle(listname).items.top(5000).getPaged();
    return items;
  }

  //Get all items with paged from a list with a filter
  public async getAllListItemsPagedWithFilter(listname: string, filter: string): Promise<any> {
    var items: any = [];
    items = await sp.web.lists.getByTitle(listname).items.top(5000).filter(filter).getPaged();
    return items;
  }

  //Get list items based on a filter and with specific columns
  public async getFilteredListItemsWithSpecificColumns(listname: string, columns: string, filter: string): Promise<any> {
    var items: any[] = [];
    items = await sp.web.lists.getByTitle(listname).items.select(columns).filter(filter).getAll();
    return items;
  }

  //Get Top list items with specific columns and sort by order
  public async getTopSortedItemsWithSpecificColumns(listname: string, columns: string, topVal: number, descColumn: string, ascColumn: string): Promise<any> {
    var items: any[] = [];
    items = await sp.web.lists.getByTitle(listname).items.select(columns).top(topVal).orderBy(descColumn, false).orderBy(ascColumn, true)();
    return items;
  }

  //Get Top list items filtered with specific columns and sort by order
  public async getFilteredTopSortedItemsWithSpecificColumns(listname: string, filter: string, columns: string, topVal: number, descColumn: string, ascColumn: string): Promise<any> {
    var items: any[] = [];
    items = await sp.web.lists.getByTitle(listname).items.select(columns).filter(filter).top(topVal).orderBy(descColumn, false).orderBy(ascColumn, true)();
    return items;
  }

  //Get choices from Choice column in a SharePoint list
  public async getChoicesFromListColumn(listname: string, columnname: string): Promise<any> {
    var choices: any = [];
    choices = await sp.web.lists.getByTitle(listname).fields.getByInternalNameOrTitle(columnname).select('Choices').get();
    return choices.Choices;
  }


  //Delete all items in a SP list
  public async deleteListItems(listname: string): Promise<any> {
    var list = sp.web.lists.getByTitle(listname);
    list.items.getAll().then((items) => {
      items.forEach(i => {
        list.items.getById(i["ID"]).delete();
      });
    });
    return true;
  }

  //Create list item
  public async createListItem(listname: string, data: any): Promise<any> {
    return sp.web.lists.getByTitle(listname).items.add(data);
  }

  //Update list item
  public updateListItem(listName: string, data: any, id: string): Promise<any> {
    return sp.web.lists.getByTitle(listName).items.getById(parseInt(id)).update(data).then(i => {
      return true;
    });
  }

  //Update multiple items with different values
  public async updateMultipleItemsWithDifferentValues(listname: string, data: any): Promise<any> {
    return new Promise<any>(async (resolve, reject) => {

      try {
        //Create object for batch
        const batch = sp.web.createBatch();
        //Get list context
        const list = await sp.web.lists.getByTitle(listname);
        const items = list.items.inBatch(batch);

        for (let itemCount = 0; itemCount < data.length; itemCount++) {
          items.getById(parseInt(data[itemCount].id)).inBatch(batch).update(data[itemCount].value);
        }
        await batch.execute().then(() => {
          resolve(true);
        }).catch((error) => {
          console.error("CommonServices_updateMultipleItems \n", error);
          reject(false);
        });
      }
      catch (error) {
        console.error("CommonServices_updateMultipleItems \n", error);
        reject(false);
      }

    });
  }

  //Update multiple items with same value
  public async updateMultipleItems(listname: string, data: any, arrayOfIds: any): Promise<any> {
    return new Promise<any>(async (resolve, reject) => {

      try {
        //Create object for batch
        const batch = sp.web.createBatch();
        //Get list context
        const list = await sp.web.lists.getByTitle(listname);
        const items = list.items.inBatch(batch);

        for (let itemCount = 0; itemCount < arrayOfIds.length; itemCount++) {
          items.getById(parseInt(arrayOfIds[itemCount])).inBatch(batch).update(data);
        }
        await batch.execute().then(() => {
          resolve(true);
        }).catch((error) => {
          console.error("CommonServices_updateMultipleItems \n", error);
          reject(false);
        });
      }
      catch (error) {
        console.error("CommonServices_updateMultipleItems \n", error);
        reject(false);
      }

    });
  }

  //Create SharePoint Site List
  public async createSPlist(listName: string) {
    try {
      const listResponse = await sp.web.lists.add(listName);
      console.log("Created list successfully. ", listName);
      return listResponse;
    }
    catch (error: any) {
      console.log("Error in creating list. ", error);
      return error;
    }
  }

  //create fields in SP lists
  public async createListFields(listname: string, fieldsToCreate: any): Promise<any> {
    return new Promise<any>(async (resolve, reject) => {
      try {
        //get list context
        const listContext = await sp.web.lists.getByTitle(listname);
        // add all the fields in a single batch call
        const batch = sp.web.createBatch();
        for (let i = 0; i < fieldsToCreate.length; i++) {
          listContext.fields.inBatch(batch).createFieldAsXml(fieldsToCreate[i]);
        }

        //execute the batch and add field to default view
        batch.execute().then(async () => {
          let addingStatus = [];
          for (let i = 0; i < fieldsToCreate.length; i++) {
            const parser = new DOMParser();
            const xml = parser.parseFromString(fieldsToCreate[i], 'text/xml');
            let fieldDisplayName = xml.querySelector('Field').getAttribute('DisplayName');
            let listView = await listContext.defaultView.fields.add(fieldDisplayName);
            addingStatus.push(listView);
          }
          Promise.all(addingStatus).then(() => {
            resolve("Success");
          });
        });
      } catch (error) {
        console.error("CommonServices_createListFields_FailedToCreatedField \n", error);
        reject("Failed");
      }
    });
  }

  //get active tournament from master list
  public async getActiveTournamentDetails(): Promise<any> {
    let filterActive: string = "Status eq '" + stringsConstants.TournamentStatusActive + "'";
    const activeTournamentsArray: any[] = await this.getItemsWithOnlyFilter(stringsConstants.TournamentsMasterList, filterActive);
    return activeTournamentsArray;
  }

  //get display names of list columns based on thier internal names
  public async getColumnsDisplayNames(listName: string, filter: string): Promise<any> {
    const columnsDisplayNames = await sp.web.lists.getByTitle(listName).fields.filter(filter).select("InternalName", "Title").get();
    return columnsDisplayNames;
  }

  //Filter and get all badge imagesfrom 'Digital Badges' library for the current user
  public async getAllBadgeImages(listName: string, userEmail: string): Promise<any> {
    try {
      let badgeImagesArray: any[] = [];
      let finalImagesArray: any[] = [];
      //If TOT is not enabled 'Tournament' column will be missing
      const filterFields = await sp.web.lists.getByTitle(stringsConstants.DigitalBadgeLibrary).fields
        .filter("Title eq 'Tournament'")
        .get();

      //if TOT is not enabled get all the badges
      if (filterFields.length == 0) {
        badgeImagesArray = await sp.web.lists.getByTitle(stringsConstants.DigitalBadgeLibrary).items.select("Title", "MinimumPoints", "File/Name").expand("File").get();
        for (let i = 0; i < badgeImagesArray.length; i++) {
          //For global badges do not check for Tournaments completion status
          finalImagesArray.push({
            title: badgeImagesArray[i].Title,
            url: rootSiteURL + "/" + listName + "/" + badgeImagesArray[i].File.Name,
            minimumPoints: badgeImagesArray[i].MinimumPoints
          });
        }
      }
      //if TOT is enabled get all the badges and filter for completed tournaments
      else {
        badgeImagesArray = await sp.web.lists.getByTitle(stringsConstants.DigitalBadgeLibrary).items.select("Title", "MinimumPoints", "Tournament/Title", "File/Name").expand("Tournament", "File").get();
        //Checking if the user is in member list
        let filterQuery = "Title eq '" + userEmail.toLowerCase() + "'" + " and Status eq 'Approved'";
        let isApprovedChampion = await this.getItemsWithOnlyFilter(stringsConstants.MemberList, filterQuery);

        //Loop through badges and filter based on user's tournaments completion status
        for (let i = 0; i < badgeImagesArray.length; i++) {
          //For global badges do not check for Tournaments completion status

          if (badgeImagesArray[i].Tournament == undefined) {
            // Show the global badges only for champion
            if (isApprovedChampion.length > 0) {
              finalImagesArray.push({
                title: badgeImagesArray[i].Title,
                url: rootSiteURL + "/" + listName + "/" + badgeImagesArray[i].File.Name,
                minimumPoints: badgeImagesArray[i].MinimumPoints
              });
            }
          }
          else if (badgeImagesArray[i].Tournament.Title !== null) {
            var tournamentCompleted = await this.getTournamentCompletedFlag(badgeImagesArray[i].Tournament.Title, userEmail);
            if (tournamentCompleted)
              finalImagesArray.push({
                title: badgeImagesArray[i].Title,
                url: rootSiteURL + "/" + listName + "/" + badgeImagesArray[i].File.Name,
                minimumPoints: 0
              });
          }
        }
      }

      return finalImagesArray;
    }
    catch (error) {
      console.error("CommonServices_getAllBadgeImages\n", error);
    }
  }
  //Check if the user has completed the tournament
  public async getTournamentCompletedFlag(tournamentName: string, currentUserEmail: string): Promise<any> {
    let tournamentCompletedFlag: boolean = false;

    tournamentName = tournamentName.replace(/'/g, "''");
    //Get total number of actions for the tournament
    let filterTournamentActions: string = "Title eq '" + tournamentName + "'";
    var tournamentActionsCount: number = 0;
    tournamentActionsCount = (await sp.web.lists.getByTitle(stringsConstants.TournamentActionsMasterList).items.filter(filterTournamentActions).select("Title").getAll()).length;

    //Get user completed actions
    let filterUserActions: string = "Tournament_x0020_Name eq '" + tournamentName + "'" + " and Title eq '" + currentUserEmail + "'";
    var userActionsCount: number = 0;
    userActionsCount = (await sp.web.lists.getByTitle(stringsConstants.UserActionsList).items.filter(filterUserActions).select("Title").getAll()).length;
    //Check if the total actions count matches with user's completed actions
    if (tournamentActionsCount == userActionsCount)
      tournamentCompletedFlag = true;
    return tournamentCompletedFlag;
  }

  //get all user action for active tournament and bind to table
  public async getUserActions(activeTournamentName: any): Promise<any> {
    return new Promise<any>(async (resolve, reject) => {
      try {
        let userActionsWithDisplayName: any = [];
        let userRanks: any = [];
        let getAllUserActions: any = [];

        //get active tournament's participants
        let filterQuery = "Tournament_x0020_Name eq '" + activeTournamentName.replace(/'/g, "''") + "'";
        //get first batch of items
        let getUserActions = await sp.web.lists.getByTitle(stringsConstants.UserActionsList).items.
          filter(filterQuery).select("Title", "Points", "UserName").top(5000).getPaged();
        if (getUserActions.results.length > 0) {
          getAllUserActions.push(...getUserActions.results);
          //get next batch, if more items found
          while (getUserActions.hasNext) {
            getUserActions = await getUserActions.getNext();
            getAllUserActions.push(...getUserActions.results);
          }
          //groupby user and sum the points
          var groupOfUniqueUsers: any = [];
          getAllUserActions.reduce((res: any, value: any) => {
            if (!res[value.Title]) {
              res[value.Title] = { Title: value.Title, Points: 0 };
              groupOfUniqueUsers.push(res[value.Title]);
            }
            res[value.Title].Points += value.Points;
            res[value.Title].UserName = value.UserName;
            return res;
          }, {});
          //sorting by points and then by display name 
          groupOfUniqueUsers.sort((a: any, b: any) => {
            if (a.Points < b.Points) return 1;
            if (a.Points > b.Points) return -1;
            if (a.Title > b.Title) return 1;
            if (a.Title < b.Title) return -1;
          });
          // get user Display Name for users
          if (groupOfUniqueUsers.length > 0) {
            for (let i = 0; i < groupOfUniqueUsers.length; i++) {
              let itemEmail: string = groupOfUniqueUsers[i].Title.toLowerCase();
              let userDisplayName = itemEmail;
              if (groupOfUniqueUsers[i].UserName != null)
                userDisplayName = groupOfUniqueUsers[i].UserName;

              if (userDisplayName.length > 0) {
                userActionsWithDisplayName.push(
                  {
                    User: userDisplayName.replace(',', ''),
                    Points: groupOfUniqueUsers[i].Points,
                    Email: itemEmail
                  });
              }
            }//for loop of getting user display name ends here   
          }
          //associate rank on the sorted array of users
          for (let j = 0; j < userActionsWithDisplayName.length; j++) {
            userRanks.push(
              {
                Rank: j + 1,
                User: userActionsWithDisplayName[j].User,
                Points: userActionsWithDisplayName[j].Points,
                Email: userActionsWithDisplayName[j].Email
              });
          }
          resolve(userRanks);
        }
        else {
          resolve(userRanks);
        }

      }
      catch (error) {
        console.error("CommonServices_getUserActions \n", error);
        reject("Failed");
      }
    });
  }

  //Get data from Tournament Actions and User Actions List for the tournament to calculate the required metrics for Tournament Report.
  public async updateCompletedTournamentDetails(completedTournamentName: string, currentDate?: Date): Promise<any> {
    return new Promise<any>(async (resolve, reject) => {
      try {

        let allUserActionsArray: any = [];
        let createParticipantReportItems: any = [];
        let columns = "Title, Points, Author/Title";
        let totalTournamentActivities: number = 0;
        let totalTournamentPoints: number = 0;
        let totalTournamentParticipants: number = 0;
        let totalCompletedParticipants: number = 0;
        let totalCompletionPercentage: number = 0;

        let filterCondition = "Title eq '" + completedTournamentName.replace(/'/g, "''") + "'";
        let filterQuery = "Tournament_x0020_Name eq '" + completedTournamentName.replace(/'/g, "''") + "'";


        //Get count of actions and sum of points for the tournament from Tournament Actions List. 
        let completedTournamentDetails: any = await sp.web.lists.getByTitle(stringsConstants.TournamentActionsMasterList).items.filter(filterCondition).getAll();
        if (completedTournamentDetails.length > 0) {
          totalTournamentActivities = completedTournamentDetails.length;
          totalTournamentPoints = completedTournamentDetails.reduce((previousValue: any, currentValue: any) => { return previousValue + currentValue["Points"]; }, 0);
        }

        //Get first batch of items from User Actions list for the tournament
        let userActionsArray = await sp.web.lists.getByTitle(stringsConstants.UserActionsList).items.filter(filterQuery).select(columns).expand("Author/Title").top(5000).getPaged();
        if (userActionsArray.results.length > 0) {
          allUserActionsArray.push(...userActionsArray.results);
          //Get next batch, if more items found in User Actions list for the tournament
          while (userActionsArray.hasNext) {
            userActionsArray = await userActionsArray.getNext();
            allUserActionsArray.push(...userActionsArray.results);
          }

          //Group the items by participants
          let organizedParticipants = this.groupBy(allUserActionsArray, (item: any) => item.Title);

          //Calculate the metrics for each participant and create an item in the Participants Report List
          organizedParticipants.forEach(async (participant) => {

            let participantName: string = participant[0].Author.Title;
            let activitiesCompleted: number = participant.length;
            let pointsCompleted: number = participant.reduce((previousValue: any, currentValue: any) => { return previousValue + currentValue["Points"]; }, 0);
            let percentageCompletion: number = Math.round((activitiesCompleted * 100) / totalTournamentActivities);

            let participantReportObject: any = this.createParticipantReportObject(completedTournamentName.trim(),
              participantName, activitiesCompleted, pointsCompleted, percentageCompletion);

            //Create an item in Participants Report List for each participant of the tournament
            this.createListItem(stringsConstants.ParticipantsReportList, participantReportObject);

            //Push the metrics of each participant into an array to calculate the total metrics for tournament.
            createParticipantReportItems.push(participantReportObject);
          });
        }
        //Calculate the total metrics related to participant details for the tournament
        if (createParticipantReportItems.length > 0) {
          totalTournamentParticipants = createParticipantReportItems.length;

          const completedParticipants = createParticipantReportItems.filter((obj: any) => {
            return obj.Completion_x0020_Percentage === 100;
          });

          if (completedParticipants.length > 0)
            totalCompletedParticipants = completedParticipants.length;

          totalCompletionPercentage = Math.round(totalCompletedParticipants * 100 / totalTournamentParticipants);
        }

        let tournamentReportObject = this.createTournamentReportObject(completedTournamentName.trim(),
          totalTournamentActivities, totalTournamentPoints, totalTournamentParticipants, totalCompletedParticipants,
          totalCompletionPercentage, currentDate);

        // Check if an item already exists in Tournaments Report list and create it if item does not exists.  
        const tournamentItem: any[] = await this.getItemsWithOnlyFilter(
          stringsConstants.TournamentsReportList, filterCondition);
        if (tournamentItem.length == 0) {
          // Create an item for the tournament in Tournaments Report list
          await this.createListItem(stringsConstants.TournamentsReportList, tournamentReportObject);
        }
        resolve(true);
      }
      catch (error) {
        console.error("CommonServices_updateCompletedTournamentDetails \n", error);
        reject(false);
      }
    });
  }

  //Get Total Points for a Member from Event Track Details list
  public async getTotalPointsForMember(memberEmailId: string): Promise<any> {

    let filterQuery = "Title eq '" + memberEmailId.toLowerCase() + "'";
    let totalPoints = 0;

    await this.getFilteredListItemsWithSpecificColumns(stringsConstants.MemberList, "ID", filterQuery)
      .then(async (memberID) => {
        //If current user is not a member skip the points calculation    
        if (memberID.length != 0) {
          let filter = "MemberId eq '" + memberID[0].ID + "'" + " and Status ne 'Pending' and Status ne 'Rejected'";
          let memberPointsArray = await this.getFilteredListItemsWithSpecificColumns(stringsConstants.EventTrackDetailsList, stringsConstants.CountColumn, filter);

          if (memberPointsArray.length > 0) {
            totalPoints = memberPointsArray.reduce((previousValue: any, currentValue: any) => { return previousValue + currentValue[stringsConstants.CountColumn]; }, 0);
          }
        }
      });

    return totalPoints;
  }

  private createParticipantReportObject = (
    tournamentName: string,
    participantName: string,
    activitiesCompleted: number,
    completedPoints: number,
    percentageCompletion: number) => {

    return {
      Title: tournamentName,
      User_x0020_Name: participantName,
      Activities_x0020_Completed: activitiesCompleted,
      Points: completedPoints,
      Completion_x0020_Percentage: percentageCompletion
    };
  }

  private createTournamentReportObject = (
    tournamentName: string,
    totalActivities: number,
    totalPoints: number,
    totalParticipants: number,
    completedParticipants: number,
    percentageCompletion: number,
    completedOn: Date) => {

    return {
      Title: tournamentName,
      Total_x0020_Activities: totalActivities,
      Total_x0020_Points: totalPoints,
      Total_x0020_Participants: totalParticipants,
      Completed_x0020_Participants: completedParticipants,
      Completion_x0020_Percentage: percentageCompletion,
      Completed_x0020_On: completedOn
    };
  }
  // Group the array of objects based on the key
  public groupBy = (list: any, keyGetter: any) => {
    const map = new Map();
    list.forEach((item: any) => {
      const key = keyGetter(item);
      const collection = map.get(key);
      if (!collection) {
        map.set(key, [item]);
      } else {
        collection.push(item);
      }
    });
    return map;
  }

  // Format Date to MMM DD, YYYY
  public formatDate = (date: any) => {
    var utc = date.toUTCString(); // 'ddd, DD MMM YYYY HH:mm:ss GMT'
    return utc.slice(8, 12) + utc.slice(5, 7) + ", " + utc.slice(12, 16);
  }
}
