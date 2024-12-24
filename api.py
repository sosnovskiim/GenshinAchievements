import asyncio

import ambr


class API:
    def __init__(self):
        self._achievement_categories = asyncio.run(
            self.__get_achievement_categories())

    def get_achievements_by_category(self, category: str):
        return self._achievement_categories[category].achievements

    @staticmethod
    async def __get_achievement_categories():
        async with ambr.AmbrAPI(lang=ambr.Language.RU) as api:
            temp = await api.fetch_achievement_categories()
            result = dict()
            element: ambr.AchievementCategory
            for element in temp:
                result[element.name] = element
            return result
