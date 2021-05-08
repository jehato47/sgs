from rest_framework import serializers
from .models import *


class SgsSerializer(serializers.ModelSerializer):
    class Meta:
        model = Sgs
        fields = '__all__'

    # def create(self, validated_data):
    #     obj = Sgs(**validated_data)
    #
    #     return obj
