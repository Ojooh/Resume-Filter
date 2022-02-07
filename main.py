from resume_filter import ResumeFilter


skills = [
    {"skill" : "Organization", "score" : "5"},
    {"skill" : "Communication", "score" : "4"},
    {"skill" : "Teamwork", "score" : "3"},
    {"skill" : "Customer service", "score" : "2"},
    {"skill" : "Responsible", "score" : "1"},
]

rf = ResumeFilter("exampley@maily.com", "passy", subject=["Administrative", "Admin", "Admin Officer"], skills=skills, server="mail.zkyte.com.ng", reply='n')
rf.filterResume()