package com.export_excel.service;

import com.export_excel.entity.ContactMessage;
import com.export_excel.exception.ConflictException;
import com.export_excel.exception.SuccessMessages;
import com.export_excel.payload.messages.ErrorMessages;
import com.export_excel.payload.request.ContactMessageRequest;
import com.export_excel.payload.response.ContactMessageResponse;
import com.export_excel.payload.response.ResponseMessage;
import com.export_excel.repository.ContactMessageRepository;
import lombok.RequiredArgsConstructor;
import org.springframework.data.domain.Page;
import org.springframework.data.domain.PageRequest;
import org.springframework.data.domain.Pageable;
import org.springframework.data.domain.Sort;
import org.springframework.http.HttpStatus;
import org.springframework.stereotype.Service;

import java.time.LocalDate;
import java.util.Objects;

@Service
@RequiredArgsConstructor
public class ContactMessageService {
    private final ContactMessageRepository contactMessageRepository;

    public ResponseMessage<ContactMessageResponse> save(ContactMessageRequest contactMessageRequest) {

        boolean isSameMessageWithSameEmailForToday=
                contactMessageRepository.existsByEmailEqualsAndDateEquals(contactMessageRequest.getEmail(), LocalDate.now());

        if(isSameMessageWithSameEmailForToday){
            throw new ConflictException(ErrorMessages.ALREADY_SEND_A_MESSAGE_TODAY);
        }

        ContactMessage contactMessage=createContactMessage(contactMessageRequest);
        ContactMessage savedData= contactMessageRepository.save(contactMessage);

        return ResponseMessage.<ContactMessageResponse>builder().
                message(SuccessMessages.MESSAGE_CREATE).
                httpStatus(HttpStatus.CREATED).
                object(createResponse(savedData)).
                build();

    }



    private ContactMessage createContactMessage(ContactMessageRequest contactMessageRequest){

        return ContactMessage.builder()
                .messageName(contactMessageRequest.getMessageName())
                .subject(contactMessageRequest.getSubject())
                .message(contactMessageRequest.getMessage())
                .email(contactMessageRequest.getEmail())
                .date(LocalDate.now()).build();
    }

    private ContactMessageResponse createResponse(ContactMessage contactMessage){

        return ContactMessageResponse.builder().
                messageName(contactMessage.getMessageName()).
                message(contactMessage.getMessage()).
                subject(contactMessage.getSubject()).
                email(contactMessage.getEmail()).
                date(contactMessage.getDate()).build();

    }

    public Page<ContactMessageResponse> getAll(int page, int size, String sort, String type) {

        Pageable pageable= createPageableObject(page, size, sort, type);

        return contactMessageRepository.findAll(pageable).map(this::createResponse);

    }




    public Page<ContactMessageResponse> searchByEmail(String email, int page, int size, String sort, String type) {

        Pageable pageable= createPageableObject(page, size, sort, type);

        return contactMessageRepository.findByEmailEquals(email,pageable).map(this::createResponse);

    }

    public Page<ContactMessageResponse> searchBySubject(String subject, int page, int size, String sort, String type) {

        Pageable pageable= createPageableObject(page, size, sort, type);

        return contactMessageRepository.findBySubjectEquals(subject,pageable).map(this::createResponse);

    }





    private Pageable createPageableObject(int page, int size, String sort, String type){

        Pageable pageable= PageRequest.of(page,size, Sort.by(sort).ascending());

        if(Objects.equals(type,"desc")){
            pageable=PageRequest.of(page,size, Sort.by(sort).descending());
        }

        return pageable;
    }


}
