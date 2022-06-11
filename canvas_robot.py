from dataclasses import dataclass, field, asdict
import logging
from canvas_robot_model import get_api_data
import canvasapi
from rich.prompt import Prompt
from rich.progress import track


@dataclass
class Answer:  # pylint: disable=too-few-public-methods
    """canvas answer see for complete list of (valid) fields
    https://canvas.instructure.com/doc/api/quiz_questions.html#:~:text=An%20Answer-,object,-looks%20like%3A
    """
    answer_html: str
    answer_weight: int


AnswerOptions = dict[str, int]  # answer_text, answer_weigth
# complete list of params : https://canvas.instructure.com/doc/api/quiz_questions.html


@dataclass
class QuestionDTO:
    answers: list[AnswerOptions]
    question_name: str = ""
    question_type: str = 'multiple_choice_question'  # other option essay question
    question_text: str = ''
    points_possible: str = '1.0'
    correct_comments: str = ''
    incorrect_comments: str = ''
    neutral_comments:  str = ''
    correct_comments_html: str = ''
    incorrect_comments_html: str = ''
    neutral_comments_html: str = ''


@dataclass
class Stats:
    quiz_ids: list[int] = field(default_factory=list)
    question_ids: list[int] = field(default_factory=list)


class CanvasRobot(object):

    def __init__(self):
        url, key = get_api_data()
        self.canvas = canvasapi.Canvas(url, key)
    # ----------------------------------------

    def get_course(self, course_id: int):
        """"":returns canvas course by its id"""
        return self.canvas.get_course(course_id)

    def get_user(self, user_id: int):
        """get user using
        :param user_id:
        :returns user
        """
        return self.canvas.get_user(user_id)

    def create_quiz(self, course_id: int, title: str, quiz_type: str = '') -> (str, int):
        """
        :param course_id:
        :param title:
        :param quiz_type:
        :return: msg, quiz_id
        """
        course: Course = self.get_course(course_id)
        quiz = course.create_quiz(dict(title=title, quiz_type=quiz_type))
        return f"{course.name} now contains {quiz}", quiz.id

    def create_question(self,
                        course_id: int,
                        quiz_id: int,
                        question_dto: QuestionDTO) -> (str, int):
        """
        :param course_id:
        :param quiz_id:
        :param question_dto:
        :return: msg, quiz_question_id
        """
        course = self.get_course(course_id)
        quiz = course.get_quiz(quiz_id)
        quiz_question = quiz.create_question(question=asdict(question_dto))
        return f"{quiz} now contains {quiz_question}", quiz_question.id

    def create_quizzes_from_data(self,
                                 course_id: int,
                                 question_format="Vraag {}.",
                                 data=None
                                 ):
        """
        :param course_id: Canvas course_id: the quizzes are added to this course
        :param question_format: used to create the question name.
        They will be numbered. Should contain '{}' is placehiolder
        starting with 1
        :param data: the quizdata
        :return:
        """
        if '{}' not in question_format:
            print(f"parameter 'question_format(={question_format})' "
                  f"should contain {{}} als placeholder")
            return False

        stats = Stats()
        for quiz_name, questions in track(data):

            msg, quiz_id = self.create_quiz(course_id=course_id,
                                            title=quiz_name,
                                            quiz_type="practice_quiz")
            logging.debug(msg)
            stats.quiz_ids.append(quiz_id)
            for index, (question_text, answers) in enumerate(questions):
                answers_asdict = [asdict(answer) for answer in answers]
                question_dto = QuestionDTO(question_name=question_format.format(index + 1),
                                           question_text=question_text,
                                           answers=answers_asdict)
                msg, question_id = self.create_question(course_id=course_id,
                                                        quiz_id=quiz_id,
                                                        question_dto=question_dto)
                logging.debug(msg)
                stats.question_ids.append(question_id)

        return stats
