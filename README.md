Sub CreateAIPresentation()
    Dim pptApp As Object
    Dim pptPres As Object
    Dim pptSlide As Object
    
    ' Create PowerPoint application
    Set pptApp = CreateObject("PowerPoint.Application")
    pptApp.Visible = True
    
    ' Create new presentation
    Set pptPres = pptApp.Presentations.Add
    
    ' Add slides
    Set pptSlide = pptPres.Slides.Add(1, 12) ' Title slide
    With pptSlide
        .Shapes.Title.TextFrame.TextRange.Text = "History of Artificial Intelligence"
        .Shapes(2).TextFrame.TextRange.Text = "Your Name"
    End With
    
    ' Slide 2: Introduction to AI
    Set pptSlide = pptPres.Slides.Add(2, 12) ' Content slide
    With pptSlide
        .Shapes.Title.TextFrame.TextRange.Text = "What is Artificial Intelligence?"
        .Shapes(2).TextFrame.TextRange.Text = "Artificial Intelligence (AI) is a field of computer science that focuses on the development of intelligent machines capable of performing tasks that would typically require human intelligence. These tasks include speech recognition, problem-solving, learning, and decision-making."
    End With
    
    ' Slide 3: Early Beginnings
    Set pptSlide = pptPres.Slides.Add(3, 12) ' Content slide
    With pptSlide
        .Shapes.Title.TextFrame.TextRange.Text = "Early Beginnings"
        .Shapes(2).TextFrame.TextRange.Text = "The concept of AI can be traced back to ancient times, with early Greek myths featuring automated human-like beings. However, the modern era of AI began in the mid-20th century. In 1956, the Dartmouth Conference marked the birth of AI as a field of research."
    End With
    
    ' Slide 4: AI Winter and Resurgence
    Set pptSlide = pptPres.Slides.Add(4, 12) ' Content slide
    With pptSlide
        .Shapes.Title.TextFrame.TextRange.Text = "AI Winter and Resurgence"
        .Shapes(2).TextFrame.TextRange.Text = "In the 1970s, AI faced a period known as 'AI Winter' due to the inability to fulfill high expectations. However, in the 1990s, with advancements in computing power and algorithmic improvements, AI experienced a resurgence. This period saw the development of expert systems, neural networks, and machine learning techniques."
    End With
    
    ' Slide 5: Recent Advancements and Future Outlook
    Set pptSlide = pptPres.Slides.Add(5, 12) ' Content slide
    With pptSlide
        .Shapes.Title.TextFrame.TextRange.Text = "Recent Advancements and Future Outlook"
        .Shapes(2).TextFrame.TextRange.Text = "In recent years, AI has made significant strides in various domains, such as natural language processing, computer vision, and robotics. Applications of AI include virtual assistants, autonomous vehicles, and medical diagnosis systems. Looking ahead, AI is poised to continue transforming industries and society, with ongoing research in areas like explainable AI, ethical considerations, and AI-powered automation."
    End With
    
    ' Save the presentation
    pptPres.SaveAs "C:\Path\To\Save\AI_History.pptx"
    
    ' Clean up
    Set pptSlide = Nothing
    Set pptPres = Nothing
    Set pptApp = Nothing
    
    MsgBox "Presentation created successfully!"
End Sub

