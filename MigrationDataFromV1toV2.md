# Migrating Data from v1 to v2

The Virtual Rounding v2 solution framework changes the underlying data storage mechanisms associated with Virtual Rounding to utilize a single site/list to store the rounding meeting information in order to allow for the Power App driven interfaces to be able to connect to the underlying list.

While the migration to v2 is not necessary, it provides several large advancements in the solution's functionality, most importantly the ability for providers to invite family members into the Teams provided meetings.

In order to minimize user impact for those organizations that deployed the v1 solution, the Teams-driven v1-style interface is still available and fully supported in v2 (although the invite experience is not fully operable without the Power Apps solution in place).

In order to migrate from v1 to v2, the following steps should be conducted before the Power Apps deployment is to occur.

1. Execute the v2 "SetupSPO.ps1" script, leveraging the v2 input RunningConfig.json file.  This will provision the single Virtual Rounding site collection to be populated in the next step. Note that this input file has an extra parameter added in to support the migration, "RoomUpnSuffix", so we can migrate the room accounts from the previous solution to the new solution.  
2. Create a Locations.csv input file for the locations you want to convert to v2 (we can leave some locations in v1 and some in v2 if so desired)
3. Execute the CopyV1ListsToV2.ps1 script. This script conducts the following:
    1. Enumerates each of the Locations
    2. Connects to each list using the VirtualRoundingRoom content type
    3. Copies all the items from the v1 list to the v2 list
    4. Creates 2 views on the v2 list, pre-filtered for the Location+Sublocation parameters (one with the usual v1 JSON formatted view, and a separate one to show the invited family members for each room)
    5. Posts these 2 views to the Teams channel associated with the Sublocation
4. Import and start the v2 JoinBotToMeeting Flow.  Note this means the v1 and v2 bot "join" flows may be running in parallel; you may want to back off the schedules of both to go on alternating schedules.

Note that the v2 migration script is non-destructive, and leaves all v1 lists as-is (and tabs pointing to the v1 list in Teams). This allows you to manually confirm the script is working as expected without affecting users (beyond showing them some new tabs for a period of time).

Once you have had a chance to validate the data and Tabs migrated as expected, you can then manually delete the "old" v1 lists (not the site, as that would take the whole Team with it, unless you intend to only use the Power App front end for the solution), remove the old v1 tabs from Teams, and stop the v1 Bot Join Flow from executing.  Scripts may be developed in the near future to automate these deletion steps.