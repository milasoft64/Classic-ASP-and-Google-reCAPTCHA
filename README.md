# Classic-ASP-and-Google-reCAPTCHA
Implementing Google reCAPTCHA with Classic ASP 

This code consists of two files. One is the JSON parser and the other is the contact form. The contact form will retain the fields if an error
occurs. It does brief validation checking and simple SQL injection prevention for the fields. If the fields pass, it will then send the
contact form to you via email.

You'll need to have CDOSYS installed for your IIS page (most providers carry this with their ASP packages)

The form uses Google's reCAPTCHA v2 to ensure you're a human rather than a robot.

