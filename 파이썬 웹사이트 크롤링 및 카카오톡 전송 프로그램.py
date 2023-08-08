import requests
from bs4 import BeautifulSoup

def get_activities():
    url = 'https://www.youtube.com/watch?v=-TkoO8Z07hI'
    try:
        response = requests.get(url, verify=True)
        response.raise_for_status()  # 예외 발생 시 에러 처리를 위한 코드 추가
        soup = BeautifulSoup(response.content, 'html.parser')

        # 웹 페이지에서 필요한 정보를 추출하는 코드를 작성합니다.
        # 예를 들어, 대외활동의 제목과 링크를 추출할 수 있습니다.
        activities = []
        for activity_elem in soup.select('.activity-item'):
            title = activity_elem.select_one('.activity-title').text.strip()
            link = activity_elem.select_one('.activity-title')['href']
            activities.append({'title': title, 'link': link})

        return activities
    except requests.exceptions.RequestException as e:
        print(f"Error: {e}")
        return []

activities = get_activities()
print(activities)
