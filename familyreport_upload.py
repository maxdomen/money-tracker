from common.DatasetLoader import DatasetLoader2

__author__ = 'Max'

def upload_familyreport():
    loader=DatasetLoader2("./",False)
    loader.upload("familyreport.xls","0Ao993IR9m-38dGhFeDg5WWtSZnBDWWlHZlBZcVdvaWc")
    #loader.upload("timesheets_review/review-projects-cur.xls","0Ao993IR9m-38dDNNQXV2RWlpcElZRC1wblk5d2xISEE")

    #loader.upload("timesheets_review/review-team-prev.xls","0Ao993IR9m-38dGdXQmloNFlYdmdPZWZtX214bTg2LUE")
    #loader.upload("timesheets_review/review-projects-prev.xls","0Ao993IR9m-38dG5pVVd2b1ZONEk2d1cyNGZ0OTdCWEE")

upload_familyreport()
print "Success"