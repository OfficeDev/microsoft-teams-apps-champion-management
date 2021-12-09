import { WebPartContext } from "@microsoft/sp-webpart-base";
import { sp } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import siteconfig from "../config/siteconfig.json";
import * as stringsConstants from "../constants/strings";

interface ICommonServicesState {

}
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

  //Get list items based on only a filter
  public async getItemsWithOnlyFilter(listname: string, filterparametres: any): Promise<any> {
    var items: any[] = [];
    items = await sp.web.lists.getByTitle(listname).items.filter(filterparametres).getAll();
    return items;
  }

  //Get all items from a list
  public async getAllListItems(listname: string): Promise<any> {
    var items: any[] = [];
    items = await sp.web.lists.getByTitle(listname).items.getAll();
    return items;
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
          let addingStatus=[];
          for (let i = 0; i < fieldsToCreate.length; i++) {
            const parser = new DOMParser();
            const xml = parser.parseFromString(fieldsToCreate[i], 'text/xml');
            let fieldDisplayName = xml.querySelector('Field').getAttribute('DisplayName');
          let listView= await listContext.views.getByTitle("All Items").fields.add(fieldDisplayName);
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



  //Filter and get all badge imagesfrom 'Digital Badges' library for the current user
  public async getAllBadgeImages(listName: string, userEmail: string): Promise<any> {
    var badgeImagesArray: any[] = [];
    var finalImagesArray: any[] = [];
    //If TOT is not enabled 'Tournament' column will be missing
    const filterFields = await sp.web.lists.getByTitle(stringsConstants.DigitalBadgeLibrary).fields
      .filter("Title eq 'Tournament'")
      .get();

    //if TOT is not enabled get all the badges
    if (filterFields.length == 0) {
      badgeImagesArray = await sp.web.lists.getByTitle(stringsConstants.DigitalBadgeLibrary).items.select("Title", "File/Name").expand("File").get();
      for (let i = 0; i < badgeImagesArray.length; i++) {
        //For global badges do not check for Tournaments completion status
        finalImagesArray.push({
          title: badgeImagesArray[i].Title,
          url: rootSiteURL + "/" + listName + "/" + badgeImagesArray[i].File.Name
        });
      }
    }
    //if TOT is enabled get all the badges and filter for completed tournaments
    else {
      badgeImagesArray = await sp.web.lists.getByTitle(stringsConstants.DigitalBadgeLibrary).items.select("Title", "Tournament/Title", "File/Name").expand("Tournament", "File").get();
      //Loop through badges and filter based on user's tournaments completion status
      for (let i = 0; i < badgeImagesArray.length; i++) {
        //For global badges do not check for Tournaments completion status

        if (badgeImagesArray[i].Tournament == undefined) {
          finalImagesArray.push({
            title: badgeImagesArray[i].Title,
            url: rootSiteURL + "/" + listName + "/" + badgeImagesArray[i].File.Name
          });
        }
        else {
          var tournamentCompleted = await this.getTournamentCompletedFlag(badgeImagesArray[i].Tournament.Title, userEmail);
          if (tournamentCompleted)
            finalImagesArray.push({
              title: badgeImagesArray[i].Title,
              url: rootSiteURL + "/" + listName + "/" + badgeImagesArray[i].File.Name
            });
        }
      }
    }

    return finalImagesArray;
  }

  //Check if the user has completed the tournament
  public async getTournamentCompletedFlag(tournamentName: string, currentUserEmail: string): Promise<any> {
    let tournamentCompletedFlag: boolean = false;
    
    tournamentName = tournamentName.replace(/'/g,"''") ;
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
   public async getUserActions(activeTournamentName,allUsersDetails): Promise<any> {
    return new Promise<any>(async (resolve, reject) => {
      try {        
        let userActionsWithDisplayName: any = [];
        let userRanks: any = [];
        let getAllUserActions: any = [];   
       
          //get active tournament's participants
          let filterQuery = "Tournament_x0020_Name eq '" +  activeTournamentName.replace(/'/g,"''")  + "'";
          //get first batch of items
          let getUserActions = await sp.web.lists.getByTitle(stringsConstants.UserActionsList).items.
            filter(filterQuery).select("Title", "Points").top(5000).getPaged();
          if (getUserActions.results.length > 0) 
          {
            getAllUserActions.push(...getUserActions.results);
            //get next batch, if more items found
            while (getUserActions.hasNext) {
              getUserActions = await getUserActions.getNext();
              getAllUserActions.push(...getUserActions.results);
            }
            //groupby user and sum the points
            var groupOfUniqueUsers = [];
            getAllUserActions.reduce((res, value) => {
              if (!res[value.Title]) {
                res[value.Title] = { Title: value.Title, Points: 0 };
                groupOfUniqueUsers.push(res[value.Title]);
              }
              res[value.Title].Points += value.Points;
              return res;
            }, {});
            //sorting by points and then by display name 
            groupOfUniqueUsers.sort((a, b) => {
              if (a.Points < b.Points) return 1;
              if (a.Points > b.Points) return -1;
              if (a.Title > b.Title) return 1;
              if (a.Title < b.Title) return -1;
            });
            // get user Display Name for users
            if (groupOfUniqueUsers.length > 0) {
              for (let i = 0; i < groupOfUniqueUsers.length; i++) {
                let itemEmail: string = groupOfUniqueUsers[i].Title.toLowerCase();
                let userDisplayName = allUsersDetails.filter(
                  (user) => user.email === itemEmail
                );
                if (userDisplayName.length > 0) {
                  userActionsWithDisplayName.push(
                    {
                      User: userDisplayName[0].displayName.replace(',',''),
                      Points: groupOfUniqueUsers[i].Points,
                      Email:itemEmail
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
                  Email:userActionsWithDisplayName[j].Email
                });
            }
            resolve(userRanks);
          }
          else{
            resolve(userRanks);
          }
          
      }
      catch (error) {
        console.error("CommonServices_getUserActions \n", error);
        reject("Failed");
        }
    });
  }
}
