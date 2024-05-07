# yourapp/middleware.py

from django.utils import timezone
from django.shortcuts import redirect
from django.contrib.sessions.models import Session

class SessionTimeoutMiddleware:
    def __init__(self, get_response):
        self.get_response = get_response

    def __call__(self, request):
        print("testing")
        #session_timeout_minutes = 1
        #last_activity = request.session.get('last_activity')
        #if not last_activity or (timezone.now() - last_activity).seconds > (session_timeout_minutes * 60):
            # Clear the session
        #    request.session.flush()
        #    return redirect('bot:login')  # Change 'login' to the name/url of your login page

        # Update last activity time in the session
        # request.session['last_activity'] = timezone.now()
        response = self.get_response(request)
        return response


        '''
        print("testing")

        # Set the session timeout duration in minutes
        session_timeout_minutes = 1

        # Get the last activity time from the session
        last_activity = request.session.get('last_activity')

        # If last activity is not set or if the session has expired
        if not last_activity or (timezone.now() - last_activity).seconds > (session_timeout_minutes * 60):
            # Clear the session
            request.session.flush()
            return redirect('bot:login')  # Change 'login' to the name/url of your login page

        # Update last activity time in the session
        # request.session['last_activity'] = timezone.now()

        response = self.get_response(request)
        return response

        '''




