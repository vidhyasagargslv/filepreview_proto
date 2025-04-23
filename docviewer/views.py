from django.shortcuts import render
from rest_framework import viewsets, status
from rest_framework.response import Response
from rest_framework.decorators import action
from django.http import HttpResponse
from .models import Document
from .serializers import DocumentSerializer
from .utils import convert_docx_to_html, update_document_html,convert_pdf_to_html
import os

class DocumentViewSet(viewsets.ModelViewSet):
    queryset = Document.objects.all()
    serializer_class = DocumentSerializer
    
    def create(self, request):
        serializer = self.get_serializer(data=request.data)
        if serializer.is_valid():
            document = serializer.save()
            file_path = document.file.path
            file_ext = os.path.splitext(file_path)[1].lower()

            html_content = None
            if file_ext == '.docx':
                html_content = convert_docx_to_html(file_path)
            elif file_ext == '.pdf':
                html_content = convert_pdf_to_html(file_path)
            elif file_ext == '.html':
                with open(file_path, 'r', encoding='utf-8') as file:
                    html_content = file.read()

            if html_content is not None:
                document.html_content = html_content
                document.save()

            return Response(serializer.data, status=status.HTTP_201_CREATED)
        return Response(serializer.errors, status=status.HTTP_400_BAD_REQUEST)
    
    @action(detail=True, methods=['post'])
    def update_content(self, request, pk=None):
        """Update the HTML content for live preview"""
        document = self.get_object()
        content = request.data.get('content', '')
        is_docx = request.data.get('is_docx', False)
        
        try:
            html_content = update_document_html(content, is_docx)
            
            # Update only the HTML content in the database
            document.html_content = html_content
            document.save(update_fields=['html_content', 'updated_at'])
            
            return Response({'html_content': html_content}, status=status.HTTP_200_OK)
        except Exception as e:
            return Response({'error': str(e)}, status=status.HTTP_400_BAD_REQUEST)
