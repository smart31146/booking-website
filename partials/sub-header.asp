<div class="sub-header-section d-flex justify-content-between">
    <div class="sub-header-section__heading d-flex flex-column align-items-start">
        <span class="subject-name"><%=Ucase(ReplaceStrQuiz(subject("subject_name")))%></span>
        <span class="page-title"><%=ReplaceStrQuiz(subject("s_topic"))%></span>
    </div>

    <div class="sub-header-section__progress d-flex align-items-center">
        <div class="topic_step">
            <span>Topic</span>
            <div class="topicStepNumberCircle"><%=topicPosition%></div> of <%=totalTopics%>
        </div>
        <div class="border border-primary vertical-line"></div>
        <div class="page_step">
            <span>Page</span>
            <div class="pageStepNumberCircle"><%=subject("s_order") - previousTopicEnd%></div> of <%=totalTopicQuestions%> in this topic</>
        </div>		
    </div>
</div>

