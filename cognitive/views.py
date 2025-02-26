from rest_framework import status
from rest_framework.decorators import api_view
from rest_framework.response import Response
from rest_framework.views import APIView

from .models import User, Attempt
from .serializers import UserSerializer, TestNSISerializer, TestResultSerializer, LoginSerializer, AttemptSerializer


@api_view(['POST'])
def register_user(request):
    if request.method == 'POST':
        serializer = UserSerializer(data=request.data)
        if serializer.is_valid():
            serializer.save()
            return Response(serializer.data, status=status.HTTP_201_CREATED)
        return Response(serializer.errors, status=status.HTTP_400_BAD_REQUEST)


@api_view(['POST'])
def save_test_result(request):
    if request.method == 'POST':
        serializer = TestResultSerializer(data=request.data)
        if serializer.is_valid():
            serializer.save()
            return Response(serializer.data, status=status.HTTP_201_CREATED)
        return Response(serializer.errors, status=status.HTTP_400_BAD_REQUEST)


@api_view(['POST'])
def login_user(request):
    if request.method == 'POST':
        serializer = LoginSerializer(data=request.data)
        if serializer.is_valid():
            username = serializer.validated_data['username']
            password = serializer.validated_data['password']

            try:
                user = User.objects.get(username=username)
            except User.DoesNotExist:
                return Response({'error': 'Пользователь не найден'}, status=status.HTTP_404_NOT_FOUND)

            if user.check_password(raw_password=password):
                return Response({'message': 'Вход выполнен успешно', 'user_id': user.id})
            else:
                return Response({'error': 'Неверный пароль'}, status=status.HTTP_400_BAD_REQUEST)
        return Response(serializer.errors, status=status.HTTP_400_BAD_REQUEST)


class UpdateAttemptView(APIView):
    def post(self, request):
        user_id = request.data.get('user_id')
        test_id = request.data.get('test_id')

        if not user_id or not test_id:
            return Response(
                {"error": "user_id и test_id обязательны"},
                status=status.HTTP_400_BAD_REQUEST,
            )

        try:
            attempt = Attempt.objects.get(user_id=user_id, test_id=test_id)
            attempt.try_count += 1  # Увеличиваем количество попыток
            attempt.save()
        except Attempt.DoesNotExist:
            # Если запись не существует, создаём новую
            attempt = Attempt.objects.create(user_id=user_id, test_id=test_id, try_count=1)

        serializer = AttemptSerializer(attempt)
        return Response(serializer.data, status=status.HTTP_200_OK)


class GetAttemptView(APIView):
    def get(self, request):
        user_id = request.query_params.get('user_id')
        test_id = request.query_params.get('test_id')

        if not user_id or not test_id:
            return Response(
                {"error": "user_id и test_id обязательны"},
                status=status.HTTP_400_BAD_REQUEST,
            )

        try:
            attempt = Attempt.objects.get(user_id=user_id, test_id=test_id)
            serializer = AttemptSerializer(attempt)
            return Response(serializer.data, status=status.HTTP_200_OK)
        except Attempt.DoesNotExist:
            return Response(
                {"try_count": 0},  # Если запись не существует, возвращаем 0 попыток
                status=status.HTTP_200_OK,
            )