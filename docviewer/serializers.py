from rest_framework import serializers
from .models import Document

class DocumentSerializer(serializers.ModelSerializer):
    class Meta:
        model = Document
        fields = ['id', 'title', 'file', 'html_content', 'created_at', 'updated_at']
        read_only_fields = ['html_content', 'created_at', 'updated_at']
